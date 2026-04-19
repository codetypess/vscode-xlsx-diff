import {
    type CellSnapshot,
    type DiffCellLocation,
    type SheetComparisonKind,
    type SheetDiffModel,
    type SheetSnapshot,
    type WorkbookDiffModel,
    type WorkbookSnapshot,
} from "../model/types";

function areMergedRangesEqual(
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null
): boolean {
    if (!leftSheet || !rightSheet) {
        return false;
    }

    if (leftSheet.mergedRanges.length !== rightSheet.mergedRanges.length) {
        return false;
    }

    return leftSheet.mergedRanges.every(
        (mergedRange, index) => mergedRange === rightSheet.mergedRanges[index]
    );
}

function areCellsEqual(
    leftCell: CellSnapshot | undefined,
    rightCell: CellSnapshot | undefined
): boolean {
    if (!leftCell && !rightCell) {
        return true;
    }

    if (!leftCell || !rightCell) {
        return false;
    }

    return (
        leftCell.displayValue === rightCell.displayValue && leftCell.formula === rightCell.formula
    );
}

function createSheetKey(
    kind: SheetComparisonKind,
    leftSheetName: string | null,
    rightSheetName: string | null
): string {
    return `${kind}:${leftSheetName ?? "-"}:${rightSheetName ?? "-"}`;
}

function createSheetDiff(
    kind: SheetComparisonKind,
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null
): SheetDiffModel {
    const diffRows = new Set<number>();
    const diffCells: DiffCellLocation[] = [];

    if (!leftSheet || !rightSheet) {
        const existingSheet = leftSheet ?? rightSheet;
        for (const cell of Object.values(existingSheet?.cells ?? {})) {
            diffRows.add(cell.rowNumber);
            diffCells.push({
                key: cell.key,
                rowNumber: cell.rowNumber,
                columnNumber: cell.columnNumber,
                address: cell.address,
            });
        }
    } else {
        const keys = new Set([...Object.keys(leftSheet.cells), ...Object.keys(rightSheet.cells)]);

        for (const key of keys) {
            const leftCell = leftSheet.cells[key];
            const rightCell = rightSheet.cells[key];

            if (areCellsEqual(leftCell, rightCell)) {
                continue;
            }

            const diffCell = leftCell ?? rightCell;
            diffRows.add(diffCell!.rowNumber);
            diffCells.push({
                key: diffCell!.key,
                rowNumber: diffCell!.rowNumber,
                columnNumber: diffCell!.columnNumber,
                address: diffCell!.address,
            });
        }
    }

    diffCells.sort((left, right) => {
        if (left.rowNumber !== right.rowNumber) {
            return left.rowNumber - right.rowNumber;
        }

        return left.columnNumber - right.columnNumber;
    });

    return {
        key: createSheetKey(kind, leftSheet?.name ?? null, rightSheet?.name ?? null),
        kind,
        leftSheet,
        rightSheet,
        leftSheetName: leftSheet?.name ?? null,
        rightSheetName: rightSheet?.name ?? null,
        rowCount: Math.max(leftSheet?.rowCount ?? 0, rightSheet?.rowCount ?? 0),
        columnCount: Math.max(leftSheet?.columnCount ?? 0, rightSheet?.columnCount ?? 0),
        diffRows: [...diffRows].sort((left, right) => left - right),
        diffCells,
        diffCellCount: diffCells.length,
        mergedRangesChanged: !areMergedRangesEqual(leftSheet, rightSheet),
    };
}

export function buildWorkbookDiff(
    leftWorkbook: WorkbookSnapshot,
    rightWorkbook: WorkbookSnapshot
): WorkbookDiffModel {
    const rightByName = new Map(rightWorkbook.sheets.map((sheet) => [sheet.name, sheet] as const));
    const matchedRightNames = new Set<string>();
    const sheets: SheetDiffModel[] = [];
    const unmatchedLeft: SheetSnapshot[] = [];

    for (const leftSheet of leftWorkbook.sheets) {
        const sameNameSheet = rightByName.get(leftSheet.name);
        if (sameNameSheet) {
            sheets.push(createSheetDiff("matched", leftSheet, sameNameSheet));
            matchedRightNames.add(sameNameSheet.name);
            continue;
        }

        unmatchedLeft.push(leftSheet);
    }

    const remainingRight = rightWorkbook.sheets.filter(
        (sheet) => !matchedRightNames.has(sheet.name)
    );
    const rightBySignature = new Map<string, SheetSnapshot[]>();

    for (const rightSheet of remainingRight) {
        const bucket = rightBySignature.get(rightSheet.signature) ?? [];
        bucket.push(rightSheet);
        rightBySignature.set(rightSheet.signature, bucket);
    }

    const removedLeft: SheetSnapshot[] = [];
    const renamedRightNames = new Set<string>();

    for (const leftSheet of unmatchedLeft) {
        const bucket = rightBySignature.get(leftSheet.signature);
        const rightSheet = bucket?.shift();

        if (rightSheet) {
            sheets.push(createSheetDiff("renamed", leftSheet, rightSheet));
            renamedRightNames.add(rightSheet.name);
            continue;
        }

        removedLeft.push(leftSheet);
    }

    for (const leftSheet of removedLeft) {
        sheets.push(createSheetDiff("removed", leftSheet, null));
    }

    for (const rightSheet of remainingRight) {
        if (renamedRightNames.has(rightSheet.name)) {
            continue;
        }

        sheets.push(createSheetDiff("added", null, rightSheet));
    }

    const diffSheets = sheets.filter(
        (sheet) => sheet.kind !== "matched" || sheet.diffCellCount > 0 || sheet.mergedRangesChanged
    );

    return {
        left: leftWorkbook,
        right: rightWorkbook,
        sheets,
        totalDiffSheets: diffSheets.length,
        totalDiffRows: diffSheets.reduce((total, sheet) => total + sheet.diffRows.length, 0),
        totalDiffCells: diffSheets.reduce((total, sheet) => total + sheet.diffCellCount, 0),
    };
}
