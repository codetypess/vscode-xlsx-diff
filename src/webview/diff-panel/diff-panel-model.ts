import { createCellKey, getCellAddress, getColumnLabel } from "../../core/model/cells";
import type {
    CellDiffStatus,
    CellSnapshot,
    SheetDiffModel,
    WorkbookDiffModel,
    WorkbookSnapshot,
} from "../../core/model/types";
import { getRuntimeMessages } from "../../i18n";
import type {
    DiffPanelFileView,
    DiffPanelRenderModel,
    DiffPanelRowView,
    DiffPanelSheetTabView,
    DiffPanelSheetView,
    DiffPanelSparseCellView,
} from "./diff-panel-types";

interface MutableSparseCellView extends DiffPanelSparseCellView {}

function getUntitledSheetLabel(): string {
    return getRuntimeMessages().workbook.untitledSheet;
}

function getWorkbookTitle(workbook: WorkbookSnapshot): string {
    return workbook.titleDetail
        ? `${workbook.fileName} (${workbook.titleDetail})`
        : workbook.fileName;
}

function getWorkbookDetailFacts(workbook: WorkbookSnapshot): DiffPanelFileView["detailFacts"] {
    return (
        workbook.detailFacts?.map((fact) => ({
            label: fact.label,
            value: fact.value,
            title: fact.titleValue,
        })) ??
        (workbook.detailLabel && workbook.detailValue
            ? [
                  {
                      label: workbook.detailLabel,
                      value: workbook.detailValue,
                      title: workbook.titleDetail,
                  },
              ]
            : [])
    );
}

function getFileViewTitle(
    workbook: WorkbookSnapshot,
    detailFacts: DiffPanelFileView["detailFacts"]
): string {
    const primaryDetail = detailFacts[0];
    if (!primaryDetail) {
        return getWorkbookTitle(workbook);
    }

    return `${workbook.fileName} (${primaryDetail.value})`;
}

function formatFileSize(bytes: number): string {
    if (bytes < 1024) {
        return `${bytes} B`;
    }

    const units = ["KB", "MB", "GB"];
    let value = bytes / 1024;
    let index = 0;

    while (value >= 1024 && index < units.length - 1) {
        value /= 1024;
        index += 1;
    }

    return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[index]}`;
}

function formatModifiedTime(value: string): string {
    return new Intl.DateTimeFormat(undefined, {
        dateStyle: "medium",
        timeStyle: "short",
    }).format(new Date(value));
}

function getSheetLabel(sheet: SheetDiffModel): string {
    if (sheet.kind === "renamed") {
        return `${sheet.leftSheetName} -> ${sheet.rightSheetName}`;
    }

    return sheet.rightSheetName ?? sheet.leftSheetName ?? getUntitledSheetLabel();
}

function getSheetHasDiff(sheet: SheetDiffModel): boolean {
    return (
        sheet.kind !== "matched" ||
        sheet.diffRows.length > 0 ||
        sheet.diffCellCount > 0 ||
        sheet.mergedRangesChanged ||
        sheet.freezePaneChanged
    );
}

function getSheetDiffTone(sheet: SheetDiffModel): CellDiffStatus {
    if (sheet.kind === "added") {
        return "added";
    }

    if (sheet.kind === "removed") {
        return "removed";
    }

    return getSheetHasDiff(sheet) ? "modified" : "equal";
}

function getDiffTonePriority(status: CellDiffStatus): number {
    switch (status) {
        case "modified":
            return 3;
        case "removed":
            return 2;
        case "added":
            return 1;
        case "equal":
        default:
            return 0;
    }
}

function mergeDiffTone(current: CellDiffStatus, next: CellDiffStatus): CellDiffStatus {
    return getDiffTonePriority(next) > getDiffTonePriority(current) ? next : current;
}

function createFileView(workbook: WorkbookSnapshot): DiffPanelFileView {
    const detailFacts = getWorkbookDetailFacts(workbook);

    return {
        title: getFileViewTitle(workbook, detailFacts),
        path: workbook.filePath,
        sizeLabel: formatFileSize(workbook.fileSize),
        detailFacts: detailFacts.slice(1),
        modifiedLabel: formatModifiedTime(workbook.modifiedTime),
        isReadonly: workbook.isReadonly ?? false,
    };
}

function createSheetTabView(
    sheet: SheetDiffModel,
    activeSheetKey: string | null
): DiffPanelSheetTabView {
    return {
        key: sheet.key,
        label: getSheetLabel(sheet),
        diffRowCount: sheet.diffRows.length,
        diffCellCount: sheet.diffCellCount,
        mergedRangesChanged: sheet.mergedRangesChanged,
        freezePaneChanged: sheet.freezePaneChanged,
        hasDiff: getSheetHasDiff(sheet),
        diffTone: getSheetDiffTone(sheet),
        isActive: sheet.key === activeSheetKey,
    };
}

function createMutableCell(rowNumber: number, columnNumber: number): MutableSparseCellView {
    return {
        key: createCellKey(rowNumber, columnNumber),
        columnNumber,
        address: getCellAddress(rowNumber, columnNumber),
        status: "equal",
        diffIndex: null,
        leftPresent: false,
        rightPresent: false,
        leftValue: "",
        rightValue: "",
        leftFormula: null,
        rightFormula: null,
    };
}

function ensureCell(
    rowsByNumber: Map<number, Map<number, MutableSparseCellView>>,
    rowNumber: number,
    columnNumber: number
): MutableSparseCellView {
    let rowCells = rowsByNumber.get(rowNumber);
    if (!rowCells) {
        rowCells = new Map<number, MutableSparseCellView>();
        rowsByNumber.set(rowNumber, rowCells);
    }

    let cell = rowCells.get(columnNumber);
    if (!cell) {
        cell = createMutableCell(rowNumber, columnNumber);
        rowCells.set(columnNumber, cell);
    }

    return cell;
}

function applyLeftCell(
    rowsByNumber: Map<number, Map<number, MutableSparseCellView>>,
    rowNumber: number,
    alignedColumnNumber: number,
    cell: CellSnapshot
): void {
    const current = ensureCell(rowsByNumber, rowNumber, alignedColumnNumber);
    current.leftPresent = true;
    current.leftValue = cell.displayValue;
    current.leftFormula = cell.formula;
}

function applyRightCell(
    rowsByNumber: Map<number, Map<number, MutableSparseCellView>>,
    rowNumber: number,
    alignedColumnNumber: number,
    cell: CellSnapshot
): void {
    const current = ensureCell(rowsByNumber, rowNumber, alignedColumnNumber);
    current.rightPresent = true;
    current.rightValue = cell.displayValue;
    current.rightFormula = cell.formula;
}

function finalizeCell(cell: MutableSparseCellView): DiffPanelSparseCellView {
    let status: CellDiffStatus = "equal";

    if (cell.leftPresent && cell.rightPresent) {
        status = cell.diffIndex === null ? "equal" : "modified";
    } else if (cell.leftPresent) {
        status = "removed";
    } else if (cell.rightPresent) {
        status = "added";
    }

    return {
        ...cell,
        status,
    };
}

function createSheetView(sheet: SheetDiffModel): DiffPanelSheetView {
    const rowsByNumber = new Map<number, Map<number, MutableSparseCellView>>();
    const diffRowsSet = new Set(sheet.diffRows);
    const leftCellsByRow = new Map<number, CellSnapshot[]>();
    const rightCellsByRow = new Map<number, CellSnapshot[]>();
    const leftAlignedColumnsBySourceColumn = new Map<number, number>();
    const rightAlignedColumnsBySourceColumn = new Map<number, number>();

    for (const cell of Object.values(sheet.leftSheet?.cells ?? {})) {
        const bucket = leftCellsByRow.get(cell.rowNumber) ?? [];
        bucket.push(cell);
        leftCellsByRow.set(cell.rowNumber, bucket);
    }

    for (const cell of Object.values(sheet.rightSheet?.cells ?? {})) {
        const bucket = rightCellsByRow.get(cell.rowNumber) ?? [];
        bucket.push(cell);
        rightCellsByRow.set(cell.rowNumber, bucket);
    }

    for (const alignedColumn of sheet.alignedColumns) {
        if (alignedColumn.leftColumnNumber !== null) {
            leftAlignedColumnsBySourceColumn.set(
                alignedColumn.leftColumnNumber,
                alignedColumn.columnNumber
            );
        }

        if (alignedColumn.rightColumnNumber !== null) {
            rightAlignedColumnsBySourceColumn.set(
                alignedColumn.rightColumnNumber,
                alignedColumn.columnNumber
            );
        }
    }

    for (const alignedRow of sheet.alignedRows) {
        for (const cell of leftCellsByRow.get(alignedRow.leftRowNumber ?? -1) ?? []) {
            const alignedColumnNumber =
                leftAlignedColumnsBySourceColumn.get(cell.columnNumber) ?? null;
            if (alignedColumnNumber === null) {
                continue;
            }

            applyLeftCell(rowsByNumber, alignedRow.rowNumber, alignedColumnNumber, cell);
        }

        for (const cell of rightCellsByRow.get(alignedRow.rightRowNumber ?? -1) ?? []) {
            const alignedColumnNumber =
                rightAlignedColumnsBySourceColumn.get(cell.columnNumber) ?? null;
            if (alignedColumnNumber === null) {
                continue;
            }

            applyRightCell(rowsByNumber, alignedRow.rowNumber, alignedColumnNumber, cell);
        }
    }

    for (const diffCell of sheet.diffCells) {
        const current = ensureCell(rowsByNumber, diffCell.rowNumber, diffCell.columnNumber);
        current.key = diffCell.key;
        current.address = diffCell.address;
        current.diffIndex = diffCell.diffIndex;
    }

    const rows: DiffPanelRowView[] = sheet.alignedRows.map((alignedRow) => {
        const rowCells = rowsByNumber.get(alignedRow.rowNumber) ?? new Map();
        let diffTone: CellDiffStatus = diffRowsSet.has(alignedRow.rowNumber)
            ? getSheetDiffTone(sheet)
            : "equal";

        const cells = [...rowCells.values()]
            .sort((left, right) => left.columnNumber - right.columnNumber)
            .map((cell) => {
                const nextCell = finalizeCell(cell);
                diffTone = mergeDiffTone(diffTone, nextCell.status);
                return nextCell;
            });

        return {
            rowNumber: alignedRow.rowNumber,
            leftRowNumber: alignedRow.leftRowNumber,
            rightRowNumber: alignedRow.rightRowNumber,
            hasDiff: diffRowsSet.has(alignedRow.rowNumber),
            diffTone,
            cells,
        };
    });

    return {
        key: sheet.key,
        label: getSheetLabel(sheet),
        leftName: sheet.leftSheetName,
        rightName: sheet.rightSheetName,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        columns: sheet.alignedColumns.map((alignedColumn) => ({
            columnNumber: alignedColumn.columnNumber,
            leftColumnNumber: alignedColumn.leftColumnNumber,
            rightColumnNumber: alignedColumn.rightColumnNumber,
            leftLabel:
                alignedColumn.leftColumnNumber === null
                    ? ""
                    : getColumnLabel(alignedColumn.leftColumnNumber),
            rightLabel:
                alignedColumn.rightColumnNumber === null
                    ? ""
                    : getColumnLabel(alignedColumn.rightColumnNumber),
        })),
        rows,
        diffRows: [...sheet.diffRows],
        diffCells: sheet.diffCells.map((cell) => ({
            key: cell.key,
            rowNumber: cell.rowNumber,
            columnNumber: cell.columnNumber,
            address: cell.address,
            diffIndex: cell.diffIndex,
        })),
        diffRowCount: sheet.diffRows.length,
        diffCellCount: sheet.diffCellCount,
        mergedRangesChanged: sheet.mergedRangesChanged,
        freezePaneChanged: sheet.freezePaneChanged,
    };
}

export function createDiffPanelRenderModel(
    diff: WorkbookDiffModel,
    activeSheetKey: string | null
): DiffPanelRenderModel {
    const activeSheet =
        diff.sheets.find((sheet) => sheet.key === activeSheetKey) ?? diff.sheets[0] ?? null;

    return {
        title: `${getWorkbookTitle(diff.left)} ↔ ${getWorkbookTitle(diff.right)}`,
        leftFile: createFileView(diff.left),
        rightFile: createFileView(diff.right),
        sheets: diff.sheets.map((sheet) => createSheetTabView(sheet, activeSheet?.key ?? null)),
        activeSheet: activeSheet ? createSheetView(activeSheet) : null,
    };
}
