export class FilterColumn {
    index(): number;
    filterType(): number;
    filterSize(): number;
    filter(index: number): string;
    addFilter(value: string): FilterColumn;
    getTop10(): { value: number; top: boolean; percent: boolean };
    setTop10(value: number, top?: boolean, percent?: boolean): FilterColumn;
    getCustomFilter(): { op1: number; v1: string; op2: number; v2: string; andOp: boolean };
    setCustomFilter(op1: number, v1: string, op2?: number, v2?: string, andOp?: boolean): FilterColumn;
    clear(): FilterColumn;
}
