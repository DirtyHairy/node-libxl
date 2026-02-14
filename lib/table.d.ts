import { AutoFilter } from './auto_filter';

export class Table {
    name(): string;
    setName(name: string): Table;
    ref(): string;
    setRef(ref: string): Table;
    autoFilter(): AutoFilter;
    style(): number;
    setStyle(style: number): Table;
    showRowStripes(): boolean;
    setShowRowStripes(showRowStripes: boolean): Table;
    showColumnStripes(): boolean;
    setShowColumnStripes(showColumnStripes: boolean): Table;
    showFirstColumn(): boolean;
    setShowFirstColumn(showFirstColumn: boolean): Table;
    showLastColumn(): boolean;
    setShowLastColumn(showLastColumn: boolean): Table;
    columnSize(): number;
    columnName(index: number): string;
    setColumnName(index: number, name: string): Table;
}
