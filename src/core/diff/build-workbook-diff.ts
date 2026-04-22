import { createCellKey, getCellAddress } from "../model/cells";
import {
    type CellSnapshot,
    type DiffCellLocation,
    type DiffRowAlignment,
    type SheetComparisonKind,
    type SheetDiffModel,
    type SheetSnapshot,
    type WorkbookDiffModel,
    type WorkbookSnapshot,
} from "../model/types";

interface RowSnapshot {
    rowNumber: number;
    signature: string;
    nonEmptyCellCount: number;
    cellsByColumn: Map<number, CellSnapshot>;
}

interface ExactRowDiffOp {
    type: "equal" | "delete" | "insert";
    leftIndex?: number;
    rightIndex?: number;
}

interface RowPairing {
    left: RowSnapshot | null;
    right: RowSnapshot | null;
}

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

function buildRowSnapshots(sheet: SheetSnapshot | null): RowSnapshot[] {
    if (!sheet || sheet.rowCount <= 0) {
        return [];
    }

    const cellsByRow = new Map<number, CellSnapshot[]>();

    for (const cell of Object.values(sheet.cells)) {
        const bucket = cellsByRow.get(cell.rowNumber) ?? [];
        bucket.push(cell);
        cellsByRow.set(cell.rowNumber, bucket);
    }

    const rows: RowSnapshot[] = [];
    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
        const cells = (cellsByRow.get(rowNumber) ?? []).sort(
            (left, right) => left.columnNumber - right.columnNumber
        );
        const cellsByColumn = new Map(cells.map((cell) => [cell.columnNumber, cell] as const));
        const signature = cells
            .map(
                (cell) =>
                    `${cell.columnNumber}\u0000${cell.displayValue}\u0000${cell.formula ?? ""}`
            )
            .join("\n");

        rows.push({
            rowNumber,
            signature,
            nonEmptyCellCount: cells.length,
            cellsByColumn,
        });
    }

    return rows;
}

function buildExactRowDiff(leftRows: RowSnapshot[], rightRows: RowSnapshot[]): ExactRowDiffOp[] {
    const leftSignatures = leftRows.map((row) => row.signature);
    const rightSignatures = rightRows.map((row) => row.signature);
    const max = leftSignatures.length + rightSignatures.length;
    const offset = max + 1;
    let furthestX = new Array<number>(offset * 2 + 1).fill(0);
    const trace: number[][] = [];

    for (let distance = 0; distance <= max; distance += 1) {
        const nextFurthestX = [...furthestX];

        for (let diagonal = -distance; diagonal <= distance; diagonal += 2) {
            let x: number;

            if (
                diagonal === -distance ||
                (diagonal !== distance &&
                    furthestX[offset + diagonal - 1] < furthestX[offset + diagonal + 1])
            ) {
                x = furthestX[offset + diagonal + 1];
            } else {
                x = furthestX[offset + diagonal - 1] + 1;
            }

            let y = x - diagonal;

            while (
                x < leftSignatures.length &&
                y < rightSignatures.length &&
                leftSignatures[x] === rightSignatures[y]
            ) {
                x += 1;
                y += 1;
            }

            nextFurthestX[offset + diagonal] = x;

            if (x >= leftSignatures.length && y >= rightSignatures.length) {
                trace.push(nextFurthestX);
                return backtrackExactRowDiff(trace, leftSignatures.length, rightSignatures.length);
            }
        }

        trace.push(nextFurthestX);
        furthestX = nextFurthestX;
    }

    return [];
}

function backtrackExactRowDiff(
    trace: number[][],
    leftLength: number,
    rightLength: number
): ExactRowDiffOp[] {
    const offset = leftLength + rightLength + 1;
    const operations: ExactRowDiffOp[] = [];
    let x = leftLength;
    let y = rightLength;

    for (let distance = trace.length - 1; distance > 0; distance -= 1) {
        const previousFurthestX = trace[distance - 1]!;
        const diagonal = x - y;
        let previousDiagonal: number;

        if (
            diagonal === -distance ||
            (diagonal !== distance &&
                previousFurthestX[offset + diagonal - 1] <
                    previousFurthestX[offset + diagonal + 1])
        ) {
            previousDiagonal = diagonal + 1;
        } else {
            previousDiagonal = diagonal - 1;
        }

        const previousX = previousFurthestX[offset + previousDiagonal] ?? 0;
        const previousY = previousX - previousDiagonal;

        while (x > previousX && y > previousY) {
            operations.push({
                type: "equal",
                leftIndex: x - 1,
                rightIndex: y - 1,
            });
            x -= 1;
            y -= 1;
        }

        if (x === previousX) {
            operations.push({
                type: "insert",
                rightIndex: y - 1,
            });
            y -= 1;
        } else {
            operations.push({
                type: "delete",
                leftIndex: x - 1,
            });
            x -= 1;
        }
    }

    while (x > 0 && y > 0) {
        operations.push({
            type: "equal",
            leftIndex: x - 1,
            rightIndex: y - 1,
        });
        x -= 1;
        y -= 1;
    }

    while (x > 0) {
        operations.push({
            type: "delete",
            leftIndex: x - 1,
        });
        x -= 1;
    }

    while (y > 0) {
        operations.push({
            type: "insert",
            rightIndex: y - 1,
        });
        y -= 1;
    }

    return operations.reverse();
}

function getRowGapCost(row: RowSnapshot): number {
    return Math.max(2, row.nonEmptyCellCount * 2);
}

function getCellMismatchCost(leftCell: CellSnapshot, rightCell: CellSnapshot): number {
    const leftDisplayValue = leftCell.displayValue.trim();
    const rightDisplayValue = rightCell.displayValue.trim();
    const leftFormula = (leftCell.formula ?? "").trim();
    const rightFormula = (rightCell.formula ?? "").trim();

    if (
        (leftDisplayValue.length > 0 &&
            rightDisplayValue.length > 0 &&
            (leftDisplayValue.includes(rightDisplayValue) ||
                rightDisplayValue.includes(leftDisplayValue))) ||
        (leftFormula.length > 0 &&
            rightFormula.length > 0 &&
            (leftFormula.includes(rightFormula) || rightFormula.includes(leftFormula)))
    ) {
        return 1;
    }

    return 2;
}

function getRowPairCost(leftRow: RowSnapshot, rightRow: RowSnapshot): number {
    if (leftRow.signature === rightRow.signature) {
        return 0;
    }

    const columnNumbers = new Set<number>([
        ...leftRow.cellsByColumn.keys(),
        ...rightRow.cellsByColumn.keys(),
    ]);
    let sharedColumnCount = 0;
    let mismatchCost = 1;

    for (const columnNumber of columnNumbers) {
        const leftCell = leftRow.cellsByColumn.get(columnNumber);
        const rightCell = rightRow.cellsByColumn.get(columnNumber);

        if (leftCell && rightCell) {
            sharedColumnCount += 1;
        }

        if (areCellsEqual(leftCell, rightCell)) {
            continue;
        }

        mismatchCost += leftCell && rightCell ? getCellMismatchCost(leftCell, rightCell) : 3;
    }

    if (sharedColumnCount === 0 && leftRow.nonEmptyCellCount > 0 && rightRow.nonEmptyCellCount > 0) {
        return getRowGapCost(leftRow) + getRowGapCost(rightRow) + 1;
    }

    return mismatchCost;
}

function alignRowSegment(leftRows: RowSnapshot[], rightRows: RowSnapshot[]): RowPairing[] {
    if (leftRows.length === 0) {
        return rightRows.map((row) => ({
            left: null,
            right: row,
        }));
    }

    if (rightRows.length === 0) {
        return leftRows.map((row) => ({
            left: row,
            right: null,
        }));
    }

    const costs = Array.from({ length: leftRows.length + 1 }, () =>
        Array<number>(rightRows.length + 1).fill(0)
    );
    const steps = Array.from({ length: leftRows.length + 1 }, () =>
        Array<"pair" | "delete" | "insert" | null>(rightRows.length + 1).fill(null)
    );

    for (let leftIndex = 1; leftIndex <= leftRows.length; leftIndex += 1) {
        costs[leftIndex]![0] = costs[leftIndex - 1]![0]! + getRowGapCost(leftRows[leftIndex - 1]!);
        steps[leftIndex]![0] = "delete";
    }

    for (let rightIndex = 1; rightIndex <= rightRows.length; rightIndex += 1) {
        costs[0]![rightIndex] =
            costs[0]![rightIndex - 1]! + getRowGapCost(rightRows[rightIndex - 1]!);
        steps[0]![rightIndex] = "insert";
    }

    for (let leftIndex = 1; leftIndex <= leftRows.length; leftIndex += 1) {
        for (let rightIndex = 1; rightIndex <= rightRows.length; rightIndex += 1) {
            const pairCost =
                costs[leftIndex - 1]![rightIndex - 1]! +
                getRowPairCost(leftRows[leftIndex - 1]!, rightRows[rightIndex - 1]!);
            const deleteCost =
                costs[leftIndex - 1]![rightIndex]! + getRowGapCost(leftRows[leftIndex - 1]!);
            const insertCost =
                costs[leftIndex]![rightIndex - 1]! + getRowGapCost(rightRows[rightIndex - 1]!);

            if (pairCost < deleteCost && pairCost < insertCost) {
                costs[leftIndex]![rightIndex] = pairCost;
                steps[leftIndex]![rightIndex] = "pair";
                continue;
            }

            if (deleteCost <= insertCost) {
                costs[leftIndex]![rightIndex] = deleteCost;
                steps[leftIndex]![rightIndex] = "delete";
                continue;
            }

            costs[leftIndex]![rightIndex] = insertCost;
            steps[leftIndex]![rightIndex] = "insert";
        }
    }

    const rows: RowPairing[] = [];
    let leftIndex = leftRows.length;
    let rightIndex = rightRows.length;

    while (leftIndex > 0 || rightIndex > 0) {
        const step = steps[leftIndex]![rightIndex];

        if (step === "pair") {
            rows.push({
                left: leftRows[leftIndex - 1]!,
                right: rightRows[rightIndex - 1]!,
            });
            leftIndex -= 1;
            rightIndex -= 1;
            continue;
        }

        if (step === "delete") {
            rows.push({
                left: leftRows[leftIndex - 1]!,
                right: null,
            });
            leftIndex -= 1;
            continue;
        }

        rows.push({
            left: null,
            right: rightRows[rightIndex - 1]!,
        });
        rightIndex -= 1;
    }

    return rows.reverse();
}

function buildAlignedRows(
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null
): RowPairing[] {
    const leftRows = buildRowSnapshots(leftSheet);
    const rightRows = buildRowSnapshots(rightSheet);

    if (leftRows.length === 0) {
        return rightRows.map((row) => ({
            left: null,
            right: row,
        }));
    }

    if (rightRows.length === 0) {
        return leftRows.map((row) => ({
            left: row,
            right: null,
        }));
    }

    const exactOperations = buildExactRowDiff(leftRows, rightRows);
    const alignedRows: RowPairing[] = [];
    let pendingLeft: RowSnapshot[] = [];
    let pendingRight: RowSnapshot[] = [];

    const flushPendingRows = () => {
        if (pendingLeft.length === 0 && pendingRight.length === 0) {
            return;
        }

        alignedRows.push(...alignRowSegment(pendingLeft, pendingRight));
        pendingLeft = [];
        pendingRight = [];
    };

    for (const operation of exactOperations) {
        if (operation.type === "equal") {
            flushPendingRows();
            alignedRows.push({
                left: leftRows[operation.leftIndex!] ?? null,
                right: rightRows[operation.rightIndex!] ?? null,
            });
            continue;
        }

        if (operation.type === "delete") {
            pendingLeft.push(leftRows[operation.leftIndex!]!);
            continue;
        }

        pendingRight.push(rightRows[operation.rightIndex!]!);
    }

    flushPendingRows();

    return alignedRows;
}

function createSheetDiff(
    kind: SheetComparisonKind,
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null
): SheetDiffModel {
    const alignedRowPairs = buildAlignedRows(leftSheet, rightSheet);
    const alignedRows: DiffRowAlignment[] = [];
    const diffRows = new Set<number>();
    const diffCells: DiffCellLocation[] = [];

    alignedRowPairs.forEach((pair, index) => {
        const rowNumber = index + 1;
        alignedRows.push({
            rowNumber,
            leftRowNumber: pair.left?.rowNumber ?? null,
            rightRowNumber: pair.right?.rowNumber ?? null,
        });

        const columnNumbers = new Set<number>([
            ...Array.from(pair.left?.cellsByColumn.keys() ?? []),
            ...Array.from(pair.right?.cellsByColumn.keys() ?? []),
        ]);

        let hasRowDiff = pair.left === null || pair.right === null;

        const sortedColumnNumbers = [...columnNumbers].sort((left, right) => left - right);
        for (const columnNumber of sortedColumnNumbers) {
            const leftCell = pair.left?.cellsByColumn.get(columnNumber);
            const rightCell = pair.right?.cellsByColumn.get(columnNumber);

            if (areCellsEqual(leftCell, rightCell)) {
                continue;
            }

            hasRowDiff = true;
            diffCells.push({
                key: createCellKey(rowNumber, columnNumber),
                rowNumber,
                columnNumber,
                address: getCellAddress(rowNumber, columnNumber),
                diffIndex: -1,
            });
        }

        if (hasRowDiff) {
            diffRows.add(rowNumber);
        }
    });

    diffCells.sort((left, right) => {
        if (left.rowNumber !== right.rowNumber) {
            return left.rowNumber - right.rowNumber;
        }

        return left.columnNumber - right.columnNumber;
    });
    diffCells.forEach((cell, index) => {
        cell.diffIndex = index;
    });

    return {
        key: createSheetKey(kind, leftSheet?.name ?? null, rightSheet?.name ?? null),
        kind,
        leftSheet,
        rightSheet,
        leftSheetName: leftSheet?.name ?? null,
        rightSheetName: rightSheet?.name ?? null,
        rowCount: alignedRows.length,
        columnCount: Math.max(leftSheet?.columnCount ?? 0, rightSheet?.columnCount ?? 0),
        alignedRows,
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
        (sheet) =>
            sheet.kind !== "matched" ||
            sheet.diffRows.length > 0 ||
            sheet.diffCellCount > 0 ||
            sheet.mergedRangesChanged
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
