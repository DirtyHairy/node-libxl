import { FilterColumn } from './filter_column';

export class AutoFilter {
    getRef(): { rowFirst: number; rowLast: number; colFirst: number; colLast: number };
    setRef(rowFirst: number, rowLast: number, colFirst: number, colLast: number): AutoFilter;
    column(colId: number): FilterColumn;
    columnSize(): number;
    columnByIndex(index: number): FilterColumn;
    getSortRange(): { rowFirst: number; rowLast: number; colFirst: number; colLast: number };
    sortLevels(): number;
    getSort(index?: number): { columnIndex: number; descending: boolean };
    setSort(columnIndex: number, descending?: boolean): AutoFilter;
    addSort(columnIndex: number, descending?: boolean): AutoFilter;
}
