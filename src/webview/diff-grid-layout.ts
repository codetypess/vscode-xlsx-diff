export function getDiffRowHeaderWidth(totalRows: number): number {
    const digits = String(Math.max(totalRows, 1)).length;
    return Math.max(56, digits * 9 + 29);
}
