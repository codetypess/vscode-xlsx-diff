import { createHash } from "node:crypto";
import { stat } from "node:fs/promises";
import * as path from "node:path";
import * as vscode from "vscode";
import {
    createCellKey,
    getCellAddress,
    hasComparableCellContent,
    normalizeCellTextLineEndings,
    parseRangeAddress,
} from "../model/cells";
import { Workbook } from "./runtime";
import {
    type SheetAutoFilterSnapshot,
    type WorkbookDetailFact,
    type CellSnapshot,
    type DefinedNameSnapshot,
    type SheetFreezePaneSnapshot,
    type SheetSnapshot,
    type SheetVisibility,
    type WorkbookSnapshot,
} from "../model/types";
import { cloneCellAlignment } from "../model/alignment";
import type { AutoFilterDefinition, CellStyleDefinition } from "fastxlsx";
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
    getColumnStyleId(columnNumber: number): number | null;
    getColumnWidth(columnNumber: number): number | null;
    getRowStyleId(rowNumber: number): number | null;
    getMergedRanges(): string[];
    getFreezePane(): SheetFreezePaneSnapshot | null;
    getAutoFilterDefinition(): AutoFilterDefinition | null;
}

interface WorkbookReader {
    getSheet(sheetName: string): SheetReader;
    getSheetNames(): string[];
    getSheetVisibility(sheetName: string): SheetVisibility;
    getStyle(styleId: number): CellStyleDefinition | null;
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

    if (sheet.autoFilter) {
        hash.update(
            `autoFilter:${sheet.autoFilter.range.startRow}:${sheet.autoFilter.range.endRow}:${sheet.autoFilter.range.startColumn}:${sheet.autoFilter.range.endColumn}\n`
        );
        if (sheet.autoFilter.sort) {
            hash.update(
                `autoFilterSort:${sheet.autoFilter.sort.columnNumber}:${sheet.autoFilter.sort.direction}\n`
            );
        }
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

function createAutoFilterSnapshot(
    definition: AutoFilterDefinition | null
): SheetAutoFilterSnapshot | null {
    if (!definition) {
        return null;
    }

    const range = parseRangeAddress(definition.range);
    if (!range) {
        return null;
    }

    const sortCondition = definition.sortState?.conditions[0] ?? null;
    return {
        range,
        sort: sortCondition
            ? {
                  columnNumber: sortCondition.columnNumber,
                  direction: sortCondition.descending ? "desc" : "asc",
              }
            : null,
    };
}

function createSparseColumnWidthsSnapshot(sheet: SheetReader): Array<number | null> {
    const columnWidths: Array<number | null> = [];

    for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
        const columnWidth = sheet.getColumnWidth(columnNumber);
        if (columnWidth === null) {
            continue;
        }

        while (columnWidths.length < columnNumber - 1) {
            columnWidths.push(null);
        }

        columnWidths.push(columnWidth);
    }

    return columnWidths;
}

function loadSheetSnapshot(workbook: WorkbookReader, sheetName: string): SheetSnapshot {
    const sheet = workbook.getSheet(sheetName);
    const cells: Record<string, CellSnapshot> = {};
    const freezePane = sheet.getFreezePane();
    const autoFilter = createAutoFilterSnapshot(sheet.getAutoFilterDefinition());
    const visibility = workbook.getSheetVisibility(sheetName);
    const columnWidths = createSparseColumnWidthsSnapshot(sheet);
    const columnAlignments = Object.fromEntries(
        Array.from({ length: sheet.columnCount }, (_, index) => index + 1)
            .flatMap((columnNumber) => {
                const styleId = sheet.getColumnStyleId(columnNumber);
                const alignment = !Number.isInteger(styleId)
                    ? null
                    : cloneCellAlignment(workbook.getStyle(styleId as number)?.alignment ?? null);
                return alignment ? [[String(columnNumber), alignment] as const] : [];
            })
            .sort(
                ([leftColumnNumber], [rightColumnNumber]) =>
                    Number(leftColumnNumber) - Number(rightColumnNumber)
            )
    );
    const rowAlignments = Object.fromEntries(
        Array.from({ length: sheet.rowCount }, (_, index) => index + 1)
            .flatMap((rowNumber) => {
                const styleId = sheet.getRowStyleId(rowNumber);
                const alignment = !Number.isInteger(styleId)
                    ? null
                    : cloneCellAlignment(workbook.getStyle(styleId as number)?.alignment ?? null);
                return alignment ? [[String(rowNumber), alignment] as const] : [];
            })
            .sort(
                ([leftRowNumber], [rightRowNumber]) =>
                    Number(leftRowNumber) - Number(rightRowNumber)
            )
    );
    const cellAlignments = {} as NonNullable<SheetSnapshot["cellAlignments"]>;

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
            const displayValue = sheet.getDisplayValue(rowNumber, columnNumber);
            const formula = sheet.getFormula(rowNumber, columnNumber);
            const styleId = sheet.getStyleId(rowNumber, columnNumber);
            const cellAlignment = !Number.isInteger(styleId)
                ? null
                : cloneCellAlignment(workbook.getStyle(styleId as number)?.alignment);

            if (cellAlignment) {
                cellAlignments[createCellKey(rowNumber, columnNumber)] = cellAlignment;
            }

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
        cellAlignments,
        rowAlignments,
        columnAlignments,
        freezePane: freezePane ? { ...freezePane } : null,
        autoFilter,
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
