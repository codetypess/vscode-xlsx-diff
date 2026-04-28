import { createHash } from "node:crypto";
import { stat } from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import {
    createCellKey,
    getCellAddress,
    hasComparableCellContent,
    normalizeCellTextLineEndings,
} from "../model/cells";
import { Workbook } from "./runtime";
import {
    type WorkbookDetailFact,
    type CellSnapshot,
    type DefinedNameSnapshot,
    type SheetFreezePaneSnapshot,
    type SheetSnapshot,
    type SheetVisibility,
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
    getColumnWidth(columnNumber: number): number | null;
    getRowHeight(rowNumber: number): number | null;
    getMergedRanges(): string[];
    getFreezePane(): SheetFreezePaneSnapshot | null;
}

interface WorkbookReader {
    getSheet(sheetName: string): SheetReader;
    getSheetNames(): string[];
    getSheetVisibility(sheetName: string): SheetVisibility;
    getDefinedNames(): DefinedNameSnapshot[];
}

interface WorkbookSnapshotMetadata {
    filePath: string;
    fileName: string;
    fileSize: number;
    modifiedTime: string;
    modifiedTimeLabel?: string;
    detailLabel?: string;
    detailValue?: string;
    detailFacts?: WorkbookDetailFact[];
    titleDetail?: string;
    isReadonly?: boolean;
}

function createSheetSignature(sheet: SheetSnapshot): string {
    const hash = createHash("sha1");
    const comparableCells = Object.values(sheet.cells)
        .filter((cell) => hasComparableCellContent(cell.displayValue, cell.formula))
        .sort((left, right) => left.key.localeCompare(right.key));
    const comparableRowCount = comparableCells.reduce(
        (maxRowNumber, cell) => Math.max(maxRowNumber, cell.rowNumber),
        0
    );
    const comparableColumnCount = comparableCells.reduce(
        (maxColumnNumber, cell) => Math.max(maxColumnNumber, cell.columnNumber),
        0
    );

    hash.update(`${comparableRowCount}:${comparableColumnCount}\n`);
    hash.update(`visibility:${sheet.visibility}\n`);

    for (const mergedRange of sheet.mergedRanges) {
        hash.update(`merge:${mergedRange}\n`);
    }

    if (sheet.freezePane) {
        hash.update(
            `freeze:${sheet.freezePane.columnCount}:${sheet.freezePane.rowCount}:${sheet.freezePane.topLeftCell}:${sheet.freezePane.activePane ?? ""}\n`
        );
    }

    for (const cell of comparableCells) {
        hash.update(
            `${cell.address}\u0000${normalizeCellTextLineEndings(cell.displayValue)}\u0000${normalizeCellTextLineEndings(
                cell.formula ?? ""
            )}\n`
        );
    }

    return hash.digest("hex");
}

function loadSheetSnapshot(workbook: WorkbookReader, sheetName: string): SheetSnapshot {
    const sheet = workbook.getSheet(sheetName);
    const cells: Record<string, CellSnapshot> = {};
    const freezePane = sheet.getFreezePane();
    const visibility = workbook.getSheetVisibility(sheetName);
    const columnWidths = Array.from({ length: sheet.columnCount }, (_, index) =>
        sheet.getColumnWidth(index + 1)
    );
    const rowHeights = Object.fromEntries(
        Array.from({ length: sheet.rowCount }, (_, index) => index + 1)
            .map((rowNumber) => [String(rowNumber), sheet.getRowHeight(rowNumber)] as const)
            .filter(([, rowHeight]) => rowHeight !== null)
    );

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
        visibility,
        mergedRanges: [...sheet.getMergedRanges()].sort((left, right) => left.localeCompare(right)),
        columnWidths,
        rowHeights,
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
    const definedNames = workbook
        .getDefinedNames()
        .map((definedName) => ({
            name: definedName.name,
            scope: definedName.scope ?? null,
            value: definedName.value,
            hidden: definedName.hidden,
        }))
        .sort((left, right) => {
            const scopeComparison = (left.scope ?? "").localeCompare(right.scope ?? "");
            if (scopeComparison !== 0) {
                return scopeComparison;
            }

            const nameComparison = left.name.localeCompare(right.name);
            if (nameComparison !== 0) {
                return nameComparison;
            }

            const valueComparison = left.value.localeCompare(right.value);
            if (valueComparison !== 0) {
                return valueComparison;
            }

            return Number(left.hidden) - Number(right.hidden);
        });

    return {
        ...metadata,
        definedNames,
        sheets,
    };
}

function createEmptyWorkbookSnapshot(metadata: WorkbookSnapshotMetadata): WorkbookSnapshot {
    return {
        ...metadata,
        definedNames: [],
        sheets: [],
    };
}

function createWorkbookDetailFacts(
    resourceDetail:
        | {
              label: string;
              value: string;
              titleValue?: string;
              extraFacts?: Array<{
                  label: string;
                  value: string;
                  titleValue?: string;
              }>;
          }
        | undefined
): WorkbookDetailFact[] | undefined {
    if (!resourceDetail) {
        return undefined;
    }

    return [
        {
            label: resourceDetail.label,
            value: resourceDetail.value,
            titleValue: resourceDetail.titleValue,
        },
        ...(resourceDetail.extraFacts ?? []).map((fact) => ({
            label: fact.label,
            value: fact.value,
            titleValue: fact.titleValue,
        })),
    ];
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
    const modifiedTimeLabel = resourceStats
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
            detailFacts: createWorkbookDetailFacts(resourceDetail),
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
        detailFacts: createWorkbookDetailFacts(resourceDetail),
        titleDetail: resourceDetail?.titleValue,
        isReadonly: isWorkbookResourceReadOnly(filePathOrUri),
    });
}
