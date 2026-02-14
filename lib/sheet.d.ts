import { Format } from "./format";
import { Font } from "./font";
import { RichString } from "./rich_string";
import { AutoFilter } from "./auto_filter";
import { FormControl } from "./form_control";
import { ConditionalFormatting } from "./conditional_formatting";
import { Table } from "./table";

export class Sheet {
    // Cell type and format
    cellType(row: number, col: number): number;
    isFormula(row: number, col: number): boolean;
    cellFormat(row: number, col: number): Format;
    setCellFormat(row: number, col: number, format: Format): Sheet;

    // Read/write string
    readStr(row: number, col: number, formatRef?: { format?: Format }): string;
    readString(row: number, col: number, formatRef?: { format?: Format }): string;
    writeStr(row: number, col: number, value: string, format?: Format): Sheet;
    writeString(row: number, col: number, value: string, format?: Format): Sheet;

    // Read/write rich string
    readRichStr(row: number, col: number, formatRef?: { format?: Format }): RichString;
    writeRichStr(row: number, col: number, richString: RichString, format?: Format): Sheet;

    // Read/write number
    readNum(row: number, col: number, formatRef?: { format?: Format }): number;
    writeNum(row: number, col: number, value: number, format?: Format): Sheet;

    // Read/write boolean
    readBool(row: number, col: number, formatRef?: { format?: Format }): boolean;
    writeBool(row: number, col: number, value: boolean, format?: Format): Sheet;

    // Read/write blank
    readBlank(row: number, col: number, formatRef?: { format?: Format }): Format;
    writeBlank(row: number, col: number, format: Format): Sheet;

    // Read/write formula
    readFormula(row: number, col: number, formatRef?: { format?: Format }): string;
    writeFormula(row: number, col: number, expr: string, format?: Format): Sheet;
    writeFormulaNum(row: number, col: number, expr: string, value: number, format?: Format): Sheet;
    writeFormulaStr(row: number, col: number, expr: string, value: string, format?: Format): Sheet;
    writeFormulaBool(row: number, col: number, expr: string, value: boolean, format?: Format): Sheet;

    // Comments
    readComment(row: number, col: number): string;
    writeComment(row: number, col: number, value: string, author?: string, width?: number, height?: number): Sheet;
    removeComment(row: number, col: number): Sheet;

    // Error
    readError(row: number, col: number): number;
    writeError(row: number, col: number, error: number, format?: Format): Sheet;

    // Cell properties
    isDate(row: number, col: number): boolean;
    isRichStr(row: number, col: number): boolean;

    // Column/row dimensions
    colWidth(col: number): number;
    rowHeight(row: number): number;
    colWidthPx(col: number): number;
    rowHeightPx(row: number): number;

    // Set column/row dimensions
    setCol(first: number, last: number, width: number, format?: Format, hidden?: boolean): Sheet;
    setRow(row: number, height: number, format?: Format, hidden?: boolean): Sheet;
    setColPx(first: number, last: number, width: number, format?: Format, hidden?: boolean): Sheet;
    setRowPx(row: number, height: number, format?: Format, hidden?: boolean): Sheet;

    // Column/row format
    rowFormat(row: number): Format;
    colFormat(col: number): Format;

    // Row/col visibility
    rowHidden(row: number): boolean;
    setRowHidden(row: number, hidden: boolean): Sheet;
    colHidden(col: number): boolean;
    setColHidden(col: number, hidden: boolean): Sheet;

    // Default row height
    defaultRowHeight(): number;
    setDefaultRowHeight(height: number): Sheet;

    // Merge
    getMerge(row: number, col: number): { rowFirst: number; rowLast: number; colFirst: number; colLast: number };
    setMerge(rowFirst: number, rowLast: number, colFirst: number, colLast: number): Sheet;
    delMerge(row: number, col: number): Sheet;
    mergeSize(): number;
    merge(index: number): { rowFirst: number; rowLast: number; colFirst: number; colLast: number };
    delMergeByIndex(index: number): Sheet;

    // Pictures
    pictureSize(): number;
    getPicture(index: number): {
        bookIndex: number;
        rowTop: number;
        colLeft: number;
        rowBottom: number;
        colRight: number;
        width: number;
        height: number;
        offset_x: number;
        offset_y: number;
        linkPath?: string;
    };
    removePicture(row: number, col: number): Sheet;
    removePictureByIndex(index: number): Sheet;
    setPicture(row: number, col: number, id: number, scale?: number, offset_x?: number, offset_y?: number, pos?: number): Sheet;
    setPicture2(row: number, col: number, id: number, width?: number, height?: number, offset_x?: number, offset_y?: number, pos?: number): Sheet;

    // Page breaks
    getHorPageBreak(index: number): number;
    getHorPageBreakSize(): number;
    getVerPageBreak(index: number): number;
    getVerPageBreakSize(): number;
    setHorPageBreak(row: number, pagebreak?: boolean): Sheet;
    setVerPageBreak(col: number, pagebreak?: boolean): Sheet;

    // Split
    split(row: number, col: number): Sheet;
    splitInfo(): { row: number; col: number };

    // Grouping
    groupRows(rowFirst: number, rowLast: number, collapsed?: boolean): Sheet;
    groupCols(colFirst: number, colLast: number, collapsed?: boolean): Sheet;
    groupSummaryBelow(): boolean;
    setGroupSummaryBelow(summaryBelow: boolean): Sheet;
    groupSummaryRight(): boolean;
    setGroupSummaryRight(summaryRight: boolean): Sheet;

    // Clear
    clear(rowFirst?: number, rowLast?: number, colFirst?: number, colLast?: number): Sheet;

    // Insert/remove rows/cols (sync)
    insertRow(rowFirst: number, rowLast: number, updateNamedRanges?: boolean): Sheet;
    insertRowSync(rowFirst: number, rowLast: number, updateNamedRanges?: boolean): Sheet;
    insertRowAsync(rowFirst: number, rowLast: number, callback: (err: Error | null) => void): Sheet;
    insertRowAsync(rowFirst: number, rowLast: number, updateNamedRanges: boolean, callback: (err: Error | null) => void): Sheet;
    insertCol(colFirst: number, colLast: number, updateNamedRanges?: boolean): Sheet;
    insertColSync(colFirst: number, colLast: number, updateNamedRanges?: boolean): Sheet;
    insertColAsync(colFirst: number, colLast: number, callback: (err: Error | null) => void): Sheet;
    insertColAsync(colFirst: number, colLast: number, updateNamedRanges: boolean, callback: (err: Error | null) => void): Sheet;
    removeRow(rowFirst: number, rowLast: number, updateNamedRanges?: boolean): Sheet;
    removeRowSync(rowFirst: number, rowLast: number, updateNamedRanges?: boolean): Sheet;
    removeRowAsync(rowFirst: number, rowLast: number, callback: (err: Error | null) => void): Sheet;
    removeRowAsync(rowFirst: number, rowLast: number, updateNamedRanges: boolean, callback: (err: Error | null) => void): Sheet;
    removeCol(colFirst: number, colLast: number, updateNamedRanges?: boolean): Sheet;
    removeColSync(colFirst: number, colLast: number, updateNamedRanges?: boolean): Sheet;
    removeColAsync(colFirst: number, colLast: number, callback: (err: Error | null) => void): Sheet;
    removeColAsync(colFirst: number, colLast: number, updateNamedRanges: boolean, callback: (err: Error | null) => void): Sheet;

    // Copy cell
    copyCell(rowSrc: number, colSrc: number, rowDst: number, colDst: number): Sheet;

    // Row/col bounds
    firstRow(): number;
    lastRow(): number;
    firstCol(): number;
    lastCol(): number;
    firstFilledRow(): number;
    lastFilledRow(): number;
    firstFilledCol(): number;
    lastFilledCol(): number;

    // Gridlines
    displayGridlines(): boolean;
    setDisplayGridlines(displayGridlines?: boolean): Sheet;
    printGridlines(): boolean;
    setPrintGridlines(printGridlines?: boolean): Sheet;

    // Zoom
    zoom(): number;
    setZoom(zoom: number): Sheet;
    printZoom(): number;
    setPrintZoom(zoom: number): Sheet;

    // Print fit
    getPrintFit(): { wPages: number; hPages: number } | false;
    setPrintFit(wPages?: number, hPages?: number): Sheet;

    // Landscape/paper
    landscape(): boolean;
    setLandscape(landscape?: boolean): Sheet;
    paper(): number;
    setPaper(paper?: number): Sheet;

    // Header/footer
    header(): string;
    setHeader(header: string, margin?: number): Sheet;
    headerMargin(): number;
    footer(): string;
    setFooter(footer: string, margin?: number): Sheet;
    footerMargin(): number;

    // Centering
    hCenter(): boolean;
    setHCenter(center?: boolean): Sheet;
    vCenter(): boolean;
    setVCenter(center?: boolean): Sheet;

    // Margins
    marginLeft(): number;
    setMarginLeft(margin: number): Sheet;
    marginRight(): number;
    setMarginRight(margin: number): Sheet;
    marginTop(): number;
    setMarginTop(margin: number): Sheet;
    marginBottom(): number;
    setMarginBottom(margin: number): Sheet;

    // Print settings
    printRowCol(): boolean;
    setPrintRowCol(printRowCol?: boolean): Sheet;
    printRepeatRows(): { rowFirst: number; rowLast: number };
    setPrintRepeatRows(rowFirst: number, rowLast: number): Sheet;
    printRepeatCols(): { colFirst: number; colLast: number };
    setPrintRepeatCols(colFirst: number, colLast: number): Sheet;
    setPrintArea(rowFirst: number, rowLast: number, colFirst: number, colLast: number): Sheet;
    clearPrintRepeats(): Sheet;
    clearPrintArea(): Sheet;

    // Named ranges
    getNamedRange(
        name: string,
        scopeId?: number,
    ): { rowFirst: number; rowLast: number; colFirst: number; colLast: number; hidden: boolean };
    setNamedRange(name: string, rowFirst: number, rowLast: number, colFirst: number, colLast: number, scopeId?: number): Sheet;
    delNamedRange(name: string, scopeId?: number): Sheet;
    namedRangeSize(): number;
    namedRange(
        index: number,
    ): { name: string; rowFirst: number; rowLast: number; colFirst: number; colLast: number; scopeId: number; hidden: boolean };

    // Tables (legacy)
    getTable(
        name: string,
    ): { rowFirst: number; rowLast: number; colFirst: number; colLast: number; headerRowCount: number; totalRowCount: number };
    tableSize(): number;
    table(
        index: number,
    ): { rowFirst: number; rowLast: number; colFirst: number; colLast: number; headerRowCount: number; totalRowCount: number };

    // Tables (new API)
    addTable(name: string, rowFirst: number, rowLast: number, colFirst: number, colLast: number, hasHeaders?: boolean, tableStyle?: number): Table;
    getTableByName(name: string): Table;
    getTableByIndex(index: number): Table;

    // Sheet name/properties
    name(): string;
    setName(name: string): Sheet;
    protect(): boolean;
    setProtect(protect?: boolean): Sheet;
    rightToLeft(): boolean;
    setRightToLeft(rightToLeft?: boolean): Sheet;
    hidden(): number;
    setHidden(state?: number): Sheet;

    // View
    getTopLeftView(): { row: number; col: number };
    setTopLeftView(row: number, col: number): Sheet;

    // Address conversion
    addrToRowCol(addr: string): { row: number; col: number; rowRelative: boolean; colRelative: boolean };
    rowColToAddr(row: number, col: number, rowRelative?: boolean, colRelative?: boolean): string;

    // Hyperlinks
    hyperlinkSize(): number;
    hyperlink(
        index: number,
    ): { hyperlink: string; rowFirst: number; rowLast: number; colFirst: number; colLast: number };
    delHyperlink(index: number): Sheet;
    addHyperlink(hyperlink: string, rowFirst: number, rowLast: number, colFirst: number, colLast: number): Sheet;
    hyperlinkIndex(row: number, col: number): number;

    // Auto filter
    isAutoFilter(): boolean;
    autoFilter(): AutoFilter;
    applyFilter(): Sheet;
    removeFilter(): Sheet;

    // Auto fit
    setAutoFitArea(rowFirst?: number, colFirst?: number, rowLast?: number, colLast?: number): Sheet;

    // Tab color
    tabColor(): number;
    setTabColor(color: number): Sheet;
    setTabColor(red: number, green: number, blue: number): Sheet;
    getTabColor(): { red: number; green: number; blue: number };

    // Ignored errors
    addIgnoredError(rowFirst: number, colFirst: number, rowLast: number, colLast: number, iError: number): Sheet;

    // Data validation
    addDataValidation(
        type: number,
        op: number,
        rowFirst: number,
        rowLast: number,
        colFirst: number,
        colLast: number,
        value1: string,
        value2?: string,
        allowBlank?: boolean,
        hideDropDown?: boolean,
        showInputMessage?: boolean,
        showErrorMessage?: boolean,
        promptTitle?: string,
        prompt?: string,
        errorTitle?: string,
        error?: string,
        errorStyle?: number,
    ): Sheet;
    addDataValidationDouble(
        type: number,
        op: number,
        rowFirst: number,
        rowLast: number,
        colFirst: number,
        colLast: number,
        value1: number,
        value2: number,
        allowBlank?: boolean,
        hideDropDown?: boolean,
        showInputMessage?: boolean,
        showErrorMessage?: boolean,
        promptTitle?: string,
        prompt?: string,
        errorTitle?: string,
        error?: string,
        errorStyle?: number,
    ): Sheet;
    removeDataValidations(): Sheet;

    // Form controls
    formControlSize(): number;
    formControl(index: number): FormControl;

    // Conditional formatting
    addConditionalFormatting(rowFirst: number, rowLast: number, colFirst: number, colLast: number): ConditionalFormatting;
    conditionalFormatting(index: number): ConditionalFormatting;
    removeConditionalFormatting(index: number): Sheet;
    conditionalFormattingSize(): number;

    // Border
    setBorder(rowFirst: number, rowLast: number, colFirst: number, colLast: number, borderStyle: number, borderColor: number): Sheet;

    // Apply filter with AutoFilter object
    applyFilter2(autoFilter: AutoFilter): Sheet;

    // Active cell / selection
    getActiveCell(): { row: number; col: number };
    setActiveCell(row: number, col: number): Sheet;
    selectionRange(): string;
    addSelectionRange(sqref: string): Sheet;
    removeSelection(): Sheet;
}
