import {
    createCellKey,
    getCellAddress,
    hasComparableCellContent,
    normalizeCellTextLineEndings,
} from "../model/cells";
import {
    type CellSnapshot,
    type DiffCellLocation,
    type DiffColumnAlignment,
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

interface ColumnSnapshot {
    columnNumber: number;
    signature: string;
    nonEmptyCellCount: number;
    cellsByRow: Map<number, CellSnapshot>;
}

interface ExactDiffOp {
    type: "equal" | "delete" | "insert";
    leftIndex?: number;
    rightIndex?: number;
}

interface RowPairing {
    left: RowSnapshot | null;
    right: RowSnapshot | null;
}

interface ColumnPairing {
    left: ColumnSnapshot | null;
    right: ColumnSnapshot | null;
}

function normalizeComparisonText(value: string | null | undefined): string | null {
    if (value === null || value === undefined) {
        return null;
    }

    return normalizeCellTextLineEndings(value);
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

function areFreezePanesEqual(
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null
): boolean {
    const leftFreezePane = leftSheet?.freezePane ?? null;
    const rightFreezePane = rightSheet?.freezePane ?? null;

    if (!leftFreezePane && !rightFreezePane) {
        return true;
    }

    if (!leftFreezePane || !rightFreezePane) {
        return false;
    }

    return (
        leftFreezePane.columnCount === rightFreezePane.columnCount &&
        leftFreezePane.rowCount === rightFreezePane.rowCount &&
        leftFreezePane.topLeftCell === rightFreezePane.topLeftCell &&
        leftFreezePane.activePane === rightFreezePane.activePane
    );
}

function areSheetVisibilitiesEqual(
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null
): boolean {
    if (!leftSheet || !rightSheet) {
        return false;
    }

    return leftSheet.visibility === rightSheet.visibility;
}

function areCellsEqual(
    leftCell: CellSnapshot | undefined,
    rightCell: CellSnapshot | undefined
): boolean {
    if (leftCell && !hasComparableCellContent(leftCell.displayValue, leftCell.formula)) {
        leftCell = undefined;
    }

    if (rightCell && !hasComparableCellContent(rightCell.displayValue, rightCell.formula)) {
        rightCell = undefined;
    }

    if (!leftCell && !rightCell) {
        return true;
    }

    if (!leftCell || !rightCell) {
        return false;
    }

    return (
        normalizeComparisonText(leftCell.displayValue) ===
            normalizeComparisonText(rightCell.displayValue) &&
        normalizeComparisonText(leftCell.formula) === normalizeComparisonText(rightCell.formula)
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
    let maxComparableRowNumber = 0;

    for (const cell of Object.values(sheet.cells)) {
        if (!hasComparableCellContent(cell.displayValue, cell.formula)) {
            continue;
        }

        const bucket = cellsByRow.get(cell.rowNumber) ?? [];
        bucket.push(cell);
        cellsByRow.set(cell.rowNumber, bucket);
        maxComparableRowNumber = Math.max(maxComparableRowNumber, cell.rowNumber);
    }

    if (maxComparableRowNumber <= 0) {
        return [];
    }

    const rows: RowSnapshot[] = [];
    for (let rowNumber = 1; rowNumber <= maxComparableRowNumber; rowNumber += 1) {
        const cells = (cellsByRow.get(rowNumber) ?? []).sort(
            (left, right) => left.columnNumber - right.columnNumber
        );
        const cellsByColumn = new Map(cells.map((cell) => [cell.columnNumber, cell] as const));
        const signature = cells
            .map(
                (cell) =>
                    `${cell.columnNumber}\u0000${normalizeComparisonText(cell.displayValue) ?? ""}\u0000${
                        normalizeComparisonText(cell.formula) ?? ""
                    }`
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

function buildColumnSnapshots(
    sheet: SheetSnapshot | null,
    alignedRowNumbersBySourceRow: ReadonlyMap<number, number>
): ColumnSnapshot[] {
    if (!sheet || sheet.columnCount <= 0) {
        return [];
    }

    const cellsByColumn = new Map<number, CellSnapshot[]>();
    let maxComparableColumnNumber = 0;

    for (const cell of Object.values(sheet.cells)) {
        if (!hasComparableCellContent(cell.displayValue, cell.formula)) {
            continue;
        }

        const bucket = cellsByColumn.get(cell.columnNumber) ?? [];
        bucket.push(cell);
        cellsByColumn.set(cell.columnNumber, bucket);
        maxComparableColumnNumber = Math.max(maxComparableColumnNumber, cell.columnNumber);
    }

    if (maxComparableColumnNumber <= 0) {
        return [];
    }

    const columns: ColumnSnapshot[] = [];
    for (let columnNumber = 1; columnNumber <= maxComparableColumnNumber; columnNumber += 1) {
        const cells = (cellsByColumn.get(columnNumber) ?? []).sort((left, right) => {
            const leftRowNumber =
                alignedRowNumbersBySourceRow.get(left.rowNumber) ?? left.rowNumber;
            const rightRowNumber =
                alignedRowNumbersBySourceRow.get(right.rowNumber) ?? right.rowNumber;
            return leftRowNumber - rightRowNumber;
        });
        const cellsByRow = new Map(
            cells.map(
                (cell) =>
                    [
                        alignedRowNumbersBySourceRow.get(cell.rowNumber) ?? cell.rowNumber,
                        cell,
                    ] as const
            )
        );
        const signature = cells
            .map((cell) => {
                const rowNumber =
                    alignedRowNumbersBySourceRow.get(cell.rowNumber) ?? cell.rowNumber;
                return `${rowNumber}\u0000${normalizeComparisonText(cell.displayValue) ?? ""}\u0000${
                    normalizeComparisonText(cell.formula) ?? ""
                }`;
            })
            .join("\n");

        columns.push({
            columnNumber,
            signature,
            nonEmptyCellCount: cells.length,
            cellsByRow,
        });
    }

    return columns;
}

function buildExactDiff(leftSignatures: string[], rightSignatures: string[]): ExactDiffOp[] {
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
                return backtrackExactDiff(trace, leftSignatures.length, rightSignatures.length);
            }
        }

        trace.push(nextFurthestX);
        furthestX = nextFurthestX;
    }

    return [];
}

function buildExactRowDiff(leftRows: RowSnapshot[], rightRows: RowSnapshot[]): ExactDiffOp[] {
    return buildExactDiff(
        leftRows.map((row) => row.signature),
        rightRows.map((row) => row.signature)
    );
}

function buildExactColumnDiff(
    leftColumns: ColumnSnapshot[],
    rightColumns: ColumnSnapshot[]
): ExactDiffOp[] {
    return buildExactDiff(
        leftColumns.map((column) => column.signature),
        rightColumns.map((column) => column.signature)
    );
}

function backtrackExactDiff(
    trace: number[][],
    leftLength: number,
    rightLength: number
): ExactDiffOp[] {
    const offset = leftLength + rightLength + 1;
    const operations: ExactDiffOp[] = [];
    let x = leftLength;
    let y = rightLength;

    for (let distance = trace.length - 1; distance > 0; distance -= 1) {
        const previousFurthestX = trace[distance - 1]!;
        const diagonal = x - y;
        let previousDiagonal: number;

        if (
            diagonal === -distance ||
            (diagonal !== distance &&
                previousFurthestX[offset + diagonal - 1] < previousFurthestX[offset + diagonal + 1])
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
    const leftDisplayValue = normalizeComparisonText(leftCell.displayValue)?.trim() ?? "";
    const rightDisplayValue = normalizeComparisonText(rightCell.displayValue)?.trim() ?? "";
    const leftFormula = normalizeComparisonText(leftCell.formula)?.trim() ?? "";
    const rightFormula = normalizeComparisonText(rightCell.formula)?.trim() ?? "";

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

    if (
        sharedColumnCount === 0 &&
        leftRow.nonEmptyCellCount > 0 &&
        rightRow.nonEmptyCellCount > 0
    ) {
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

function getColumnGapCost(column: ColumnSnapshot): number {
    return Math.max(2, column.nonEmptyCellCount * 2);
}

function getColumnPairCost(leftColumn: ColumnSnapshot, rightColumn: ColumnSnapshot): number {
    if (leftColumn.signature === rightColumn.signature) {
        return 0;
    }

    const rowNumbers = new Set<number>([
        ...leftColumn.cellsByRow.keys(),
        ...rightColumn.cellsByRow.keys(),
    ]);
    let sharedRowCount = 0;
    let mismatchCost = 1;

    for (const rowNumber of rowNumbers) {
        const leftCell = leftColumn.cellsByRow.get(rowNumber);
        const rightCell = rightColumn.cellsByRow.get(rowNumber);

        if (leftCell && rightCell) {
            sharedRowCount += 1;
        }

        if (areCellsEqual(leftCell, rightCell)) {
            continue;
        }

        mismatchCost += leftCell && rightCell ? getCellMismatchCost(leftCell, rightCell) : 3;
    }

    if (
        sharedRowCount === 0 &&
        leftColumn.nonEmptyCellCount > 0 &&
        rightColumn.nonEmptyCellCount > 0
    ) {
        return getColumnGapCost(leftColumn) + getColumnGapCost(rightColumn) + 1;
    }

    return mismatchCost;
}

function alignColumnSegment(
    leftColumns: ColumnSnapshot[],
    rightColumns: ColumnSnapshot[]
): ColumnPairing[] {
    if (leftColumns.length === 0) {
        return rightColumns.map((column) => ({
            left: null,
            right: column,
        }));
    }

    if (rightColumns.length === 0) {
        return leftColumns.map((column) => ({
            left: column,
            right: null,
        }));
    }

    const costs = Array.from({ length: leftColumns.length + 1 }, () =>
        Array<number>(rightColumns.length + 1).fill(0)
    );
    const steps = Array.from({ length: leftColumns.length + 1 }, () =>
        Array<"pair" | "delete" | "insert" | null>(rightColumns.length + 1).fill(null)
    );

    for (let leftIndex = 1; leftIndex <= leftColumns.length; leftIndex += 1) {
        costs[leftIndex]![0] =
            costs[leftIndex - 1]![0]! + getColumnGapCost(leftColumns[leftIndex - 1]!);
        steps[leftIndex]![0] = "delete";
    }

    for (let rightIndex = 1; rightIndex <= rightColumns.length; rightIndex += 1) {
        costs[0]![rightIndex] =
            costs[0]![rightIndex - 1]! + getColumnGapCost(rightColumns[rightIndex - 1]!);
        steps[0]![rightIndex] = "insert";
    }

    for (let leftIndex = 1; leftIndex <= leftColumns.length; leftIndex += 1) {
        for (let rightIndex = 1; rightIndex <= rightColumns.length; rightIndex += 1) {
            const pairCost =
                costs[leftIndex - 1]![rightIndex - 1]! +
                getColumnPairCost(leftColumns[leftIndex - 1]!, rightColumns[rightIndex - 1]!);
            const deleteCost =
                costs[leftIndex - 1]![rightIndex]! + getColumnGapCost(leftColumns[leftIndex - 1]!);
            const insertCost =
                costs[leftIndex]![rightIndex - 1]! +
                getColumnGapCost(rightColumns[rightIndex - 1]!);

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

    const columns: ColumnPairing[] = [];
    let leftIndex = leftColumns.length;
    let rightIndex = rightColumns.length;

    while (leftIndex > 0 || rightIndex > 0) {
        const step = steps[leftIndex]![rightIndex];

        if (step === "pair") {
            columns.push({
                left: leftColumns[leftIndex - 1]!,
                right: rightColumns[rightIndex - 1]!,
            });
            leftIndex -= 1;
            rightIndex -= 1;
            continue;
        }

        if (step === "delete") {
            columns.push({
                left: leftColumns[leftIndex - 1]!,
                right: null,
            });
            leftIndex -= 1;
            continue;
        }

        columns.push({
            left: null,
            right: rightColumns[rightIndex - 1]!,
        });
        rightIndex -= 1;
    }

    return columns.reverse();
}

function createAlignedRowNumberMap(
    pairs: RowPairing[],
    side: "left" | "right"
): Map<number, number> {
    const alignedRowNumbersBySourceRow = new Map<number, number>();

    pairs.forEach((pair, index) => {
        const rowNumber = side === "left" ? pair.left?.rowNumber : pair.right?.rowNumber;
        if (rowNumber !== undefined && rowNumber !== null) {
            alignedRowNumbersBySourceRow.set(rowNumber, index + 1);
        }
    });

    return alignedRowNumbersBySourceRow;
}

function buildAlignedColumns(
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null,
    leftAlignedRowNumbersBySourceRow: ReadonlyMap<number, number>,
    rightAlignedRowNumbersBySourceRow: ReadonlyMap<number, number>
): ColumnPairing[] {
    const leftColumns = buildColumnSnapshots(leftSheet, leftAlignedRowNumbersBySourceRow);
    const rightColumns = buildColumnSnapshots(rightSheet, rightAlignedRowNumbersBySourceRow);

    if (leftColumns.length === 0) {
        return rightColumns.map((column) => ({
            left: null,
            right: column,
        }));
    }

    if (rightColumns.length === 0) {
        return leftColumns.map((column) => ({
            left: column,
            right: null,
        }));
    }

    const exactOperations = buildExactColumnDiff(leftColumns, rightColumns);
    const alignedColumns: ColumnPairing[] = [];
    let pendingLeft: ColumnSnapshot[] = [];
    let pendingRight: ColumnSnapshot[] = [];

    const flushPendingColumns = () => {
        if (pendingLeft.length === 0 && pendingRight.length === 0) {
            return;
        }

        alignedColumns.push(...alignColumnSegment(pendingLeft, pendingRight));
        pendingLeft = [];
        pendingRight = [];
    };

    for (const operation of exactOperations) {
        if (operation.type === "equal") {
            flushPendingColumns();
            alignedColumns.push({
                left: leftColumns[operation.leftIndex!] ?? null,
                right: rightColumns[operation.rightIndex!] ?? null,
            });
            continue;
        }

        if (operation.type === "delete") {
            pendingLeft.push(leftColumns[operation.leftIndex!]!);
            continue;
        }

        pendingRight.push(rightColumns[operation.rightIndex!]!);
    }

    flushPendingColumns();

    return alignedColumns;
}

function createSheetDiff(
    kind: SheetComparisonKind,
    leftSheet: SheetSnapshot | null,
    rightSheet: SheetSnapshot | null,
    options: {
        sheetOrderChanged?: boolean;
    } = {}
): SheetDiffModel {
    const alignedRowPairs = buildAlignedRows(leftSheet, rightSheet);
    const leftAlignedRowNumbersBySourceRow = createAlignedRowNumberMap(alignedRowPairs, "left");
    const rightAlignedRowNumbersBySourceRow = createAlignedRowNumberMap(alignedRowPairs, "right");
    const alignedColumnPairs = buildAlignedColumns(
        leftSheet,
        rightSheet,
        leftAlignedRowNumbersBySourceRow,
        rightAlignedRowNumbersBySourceRow
    );
    const alignedRows: DiffRowAlignment[] = [];
    const alignedColumns: DiffColumnAlignment[] = alignedColumnPairs.map((pair, index) => ({
        columnNumber: index + 1,
        leftColumnNumber: pair.left?.columnNumber ?? null,
        rightColumnNumber: pair.right?.columnNumber ?? null,
    }));
    const diffRows = new Set<number>();
    const diffCells: DiffCellLocation[] = [];
    const leftAlignedColumnsBySourceColumn = new Map<number, number>();
    const rightAlignedColumnsBySourceColumn = new Map<number, number>();

    alignedColumns.forEach((column) => {
        if (column.leftColumnNumber !== null) {
            leftAlignedColumnsBySourceColumn.set(column.leftColumnNumber, column.columnNumber);
        }
        if (column.rightColumnNumber !== null) {
            rightAlignedColumnsBySourceColumn.set(column.rightColumnNumber, column.columnNumber);
        }
    });

    alignedRowPairs.forEach((pair, index) => {
        const rowNumber = index + 1;
        alignedRows.push({
            rowNumber,
            leftRowNumber: pair.left?.rowNumber ?? null,
            rightRowNumber: pair.right?.rowNumber ?? null,
        });

        const alignedColumnNumbers = new Set<number>();
        for (const sourceColumnNumber of pair.left?.cellsByColumn.keys() ?? []) {
            const alignedColumnNumber =
                leftAlignedColumnsBySourceColumn.get(sourceColumnNumber) ?? null;
            if (alignedColumnNumber !== null) {
                alignedColumnNumbers.add(alignedColumnNumber);
            }
        }
        for (const sourceColumnNumber of pair.right?.cellsByColumn.keys() ?? []) {
            const alignedColumnNumber =
                rightAlignedColumnsBySourceColumn.get(sourceColumnNumber) ?? null;
            if (alignedColumnNumber !== null) {
                alignedColumnNumbers.add(alignedColumnNumber);
            }
        }

        let hasRowDiff = pair.left === null || pair.right === null;

        const sortedColumnNumbers = [...alignedColumnNumbers].sort((left, right) => left - right);
        for (const columnNumber of sortedColumnNumbers) {
            const alignedColumn = alignedColumns[columnNumber - 1];
            const leftCell =
                alignedColumn?.leftColumnNumber !== null
                    ? pair.left?.cellsByColumn.get(alignedColumn.leftColumnNumber)
                    : undefined;
            const rightCell =
                alignedColumn?.rightColumnNumber !== null
                    ? pair.right?.cellsByColumn.get(alignedColumn.rightColumnNumber)
                    : undefined;

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
        columnCount: alignedColumns.length,
        alignedRows,
        alignedColumns,
        diffRows: [...diffRows].sort((left, right) => left - right),
        diffCells,
        diffCellCount: diffCells.length,
        mergedRangesChanged: !areMergedRangesEqual(leftSheet, rightSheet),
        freezePaneChanged: !areFreezePanesEqual(leftSheet, rightSheet),
        visibilityChanged: !areSheetVisibilitiesEqual(leftSheet, rightSheet),
        sheetOrderChanged: options.sheetOrderChanged ?? false,
    };
}

export function buildWorkbookDiff(
    leftWorkbook: WorkbookSnapshot,
    rightWorkbook: WorkbookSnapshot
): WorkbookDiffModel {
    const rightByName = new Map(
        rightWorkbook.sheets.map((sheet, index) => [sheet.name, { sheet, index }] as const)
    );
    const matchedRightNames = new Set<string>();
    const sheets: SheetDiffModel[] = [];
    const unmatchedLeft: Array<{
        sheet: SheetSnapshot;
        index: number;
    }> = [];

    for (const [leftIndex, leftSheet] of leftWorkbook.sheets.entries()) {
        const sameNameEntry = rightByName.get(leftSheet.name);
        if (sameNameEntry) {
            sheets.push(
                createSheetDiff("matched", leftSheet, sameNameEntry.sheet, {
                    sheetOrderChanged: leftIndex !== sameNameEntry.index,
                })
            );
            matchedRightNames.add(sameNameEntry.sheet.name);
            continue;
        }

        unmatchedLeft.push({
            sheet: leftSheet,
            index: leftIndex,
        });
    }

    const remainingRight = rightWorkbook.sheets
        .map((sheet, index) => ({ sheet, index }))
        .filter(({ sheet }) => !matchedRightNames.has(sheet.name));
    const rightBySignature = new Map<
        string,
        Array<{
            sheet: SheetSnapshot;
            index: number;
        }>
    >();

    for (const rightEntry of remainingRight) {
        const bucket = rightBySignature.get(rightEntry.sheet.signature) ?? [];
        bucket.push(rightEntry);
        rightBySignature.set(rightEntry.sheet.signature, bucket);
    }

    const removedLeft: Array<{
        sheet: SheetSnapshot;
        index: number;
    }> = [];
    const renamedRightNames = new Set<string>();

    for (const leftEntry of unmatchedLeft) {
        const bucket = rightBySignature.get(leftEntry.sheet.signature);
        const rightEntry = bucket?.shift();

        if (rightEntry) {
            sheets.push(
                createSheetDiff("renamed", leftEntry.sheet, rightEntry.sheet, {
                    sheetOrderChanged: leftEntry.index !== rightEntry.index,
                })
            );
            renamedRightNames.add(rightEntry.sheet.name);
            continue;
        }

        removedLeft.push(leftEntry);
    }

    for (const leftEntry of removedLeft) {
        sheets.push(createSheetDiff("removed", leftEntry.sheet, null));
    }

    for (const rightEntry of remainingRight) {
        if (renamedRightNames.has(rightEntry.sheet.name)) {
            continue;
        }

        sheets.push(createSheetDiff("added", null, rightEntry.sheet));
    }

    const diffSheets = sheets.filter(
        (sheet) =>
            sheet.kind !== "matched" ||
            sheet.diffRows.length > 0 ||
            sheet.diffCellCount > 0 ||
            sheet.mergedRangesChanged ||
            sheet.freezePaneChanged ||
            sheet.visibilityChanged ||
            sheet.sheetOrderChanged
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
