/// <reference types="node" />
import { Sheet } from './sheet';
import { Format } from './format';
import { Font } from './font';
import { RichString } from './rich_string';
import { CoreProperties } from './core_properties';
import { ConditionalFormat } from './conditional_format';

export class Book {
    constructor(type: number);

    // Load from file
    loadSync(filename: string, tempfile?: string): Book;
    load(filename: string, callback: (err: Error | null) => void): Book;
    load(filename: string, tempfile: string, callback: (err: Error | null) => void): Book;
    loadAsync(filename: string, callback: (err: Error | null) => void): Book;
    loadAsync(filename: string, tempfile: string, callback: (err: Error | null) => void): Book;

    // Load specific sheet
    loadSheetSync(filename: string, sheetIndex: number, tempfile?: string, keepAllSheets?: boolean): Book;
    loadSheet(filename: string, sheetIndex: number, callback: (err: Error | null) => void): Book;
    loadSheet(filename: string, sheetIndex: number, tempfile: string, callback: (err: Error | null) => void): Book;
    loadSheet(
        filename: string,
        sheetIndex: number,
        tempfile: string,
        keepAllSheets: boolean,
        callback: (err: Error | null) => void,
    ): Book;
    loadSheetAsync(filename: string, sheetIndex: number, callback: (err: Error | null) => void): Book;
    loadSheetAsync(filename: string, sheetIndex: number, tempfile: string, callback: (err: Error | null) => void): Book;
    loadSheetAsync(
        filename: string,
        sheetIndex: number,
        tempfile: string,
        keepAllSheets: boolean,
        callback: (err: Error | null) => void,
    ): Book;

    // Load partially
    loadPartiallySync(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        tempfile?: string,
        keepAllSheets?: boolean,
    ): Book;
    loadPartially(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        callback: (err: Error | null) => void,
    ): Book;
    loadPartially(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        tempfile: string,
        callback: (err: Error | null) => void,
    ): Book;
    loadPartially(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        tempfile: string,
        keepAllSheets: boolean,
        callback: (err: Error | null) => void,
    ): Book;
    loadPartiallyAsync(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        callback: (err: Error | null) => void,
    ): Book;
    loadPartiallyAsync(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        tempfile: string,
        callback: (err: Error | null) => void,
    ): Book;
    loadPartiallyAsync(
        filename: string,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        tempfile: string,
        keepAllSheets: boolean,
        callback: (err: Error | null) => void,
    ): Book;

    // Load without empty cells
    loadWithoutEmptyCellsSync(filename: string): Book;
    loadWithoutEmptyCells(filename: string, callback: (err: Error | null) => void): Book;
    loadWithoutEmptyCellsAsync(filename: string, callback: (err: Error | null) => void): Book;

    // Load info
    loadInfoSync(filename: string): Book;
    loadInfo(filename: string, callback: (err: Error | null) => void): Book;
    loadInfoAsync(filename: string, callback: (err: Error | null) => void): Book;

    // Write/save to file
    writeSync(filename: string, useTempFile?: boolean): Book;
    saveSync(filename: string, useTempFile?: boolean): Book;
    write(filename: string, callback: (err: Error | null) => void): Book;
    write(filename: string, useTempFile: boolean, callback: (err: Error | null, result: void) => void): Book;
    save(filename: string, callback: (err: Error | null) => void): Book;
    save(filename: string, useTempFile: boolean, callback: (err: Error | null) => void): Book;
    writeAsync(filename: string, callback: (err: Error | null) => void): Book;
    writeAsync(filename: string, useTempFile: boolean, callback: (err: Error | null) => void): Book;
    saveAsync(filename: string, callback: (err: Error | null) => void): Book;
    saveAsync(filename: string, useTempFile: boolean, callback: (err: Error | null) => void): Book;

    // Load from buffer
    loadRawSync(
        buffer: Buffer,
        sheetIndex?: number,
        firstRow?: number,
        lastRow?: number,
        keepAllSheets?: boolean,
    ): Book;
    loadRaw(buffer: Buffer, callback: (err: Error | null) => void): Book;
    loadRaw(buffer: Buffer, sheetIndex: number, callback: (err: Error | null) => void): Book;
    loadRaw(buffer: Buffer, sheetIndex: number, firstRow: number, callback: (err: Error | null) => void): Book;
    loadRaw(
        buffer: Buffer,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        callback: (err: Error | null) => void,
    ): Book;
    loadRaw(
        buffer: Buffer,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        keepAllSheets: boolean,
        callback: (err: Error | null) => void,
    ): Book;
    loadRawAsync(buffer: Buffer, callback: (err: Error | null) => void): Book;
    loadRawAsync(buffer: Buffer, sheetIndex: number, callback: (err: Error | null) => void): Book;
    loadRawAsync(buffer: Buffer, sheetIndex: number, firstRow: number, callback: (err: Error | null) => void): Book;
    loadRawAsync(
        buffer: Buffer,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        callback: (err: Error | null) => void,
    ): Book;
    loadRawAsync(
        buffer: Buffer,
        sheetIndex: number,
        firstRow: number,
        lastRow: number,
        keepAllSheets: boolean,
        callback: (err: Error | null) => void,
    ): Book;

    // Write to buffer
    writeRawSync(): Buffer;
    saveRawSync(): Buffer;
    writeRaw(callback: (err: Error | null, buffer?: Buffer) => void): Book;
    writeRawAsync(callback: (err: Error | null, buffer?: Buffer) => void): Book;
    saveRaw(callback: (err: Error | null, buffer?: Buffer) => void): Book;
    saveRawAsync(callback: (err: Error | null, buffer?: Buffer) => void): Book;

    // Load info from buffer
    loadInfoRawSync(buffer: Buffer): Book;
    loadInfoRaw(buffer: Buffer, callback: (err: Error | null) => void): Book;
    loadInfoRawAsync(buffer: Buffer, callback: (err: Error | null) => void): Book;

    // Sheet management
    addSheet(name: string, parentSheet?: Sheet): Sheet;
    insertSheet(index: number, name: string, parentSheet?: Sheet): Sheet;
    getSheet(index: number): Sheet;
    getSheetName(index: number): string;
    sheetType(index: number): number;
    moveSheet(srcIndex: number, destIndex: number): Book;
    delSheet(index: number): Book;
    sheetCount(): number;

    // Format management
    addFormat(parentFormat?: Format): Format;
    addFormatFromStyle(style: number): Format;
    format(index: number): Format;
    formatSize(): number;

    // Font management
    addFont(parentFont?: Font): Font;
    font(index: number): Font;
    fontSize(): number;

    // Conditional format management
    addConditionalFormat(): ConditionalFormat;
    conditionalFormat(index: number): ConditionalFormat;
    conditionalFormatSize(): number;

    // Rich string management
    addRichString(): RichString;

    // Custom number format
    addCustomNumFormat(description: string): number;
    customNumFormat(index: number): string;

    // Date utilities
    datePack(
        year: number,
        month: number,
        day: number,
        hour?: number,
        minute?: number,
        second?: number,
        msecond?: number,
    ): number;
    dateUnpack(value: number): {
        year: number;
        month: number;
        day: number;
        hour: number;
        minute: number;
        second: number;
        msecond: number;
    };

    // Color utilities
    colorPack(red: number, green: number, blue: number): number;
    colorUnpack(value: number): { red: number; green: number; blue: number };

    // Active sheet
    activeSheet(): number;
    setActiveSheet(index: number): Book;

    // Pictures
    pictureSize(): number;
    getPicture(index: number): { type: number; data: Buffer };
    getPictureSync(index: number): { type: number; data: Buffer };
    getPictureAsync(index: number, callback: (err: Error | null, type?: number, data?: Buffer) => void): Book;
    addPicture(filename: string): number;
    addPicture(buffer: Buffer): number;
    addPictureSync(filename: string): number;
    addPictureSync(buffer: Buffer): number;
    addPictureAsync(filename: string, callback: (err: Error | null, id?: number) => void): Book;
    addPictureAsync(buffer: Buffer, callback: (err: Error | null, id?: number) => void): Book;
    addPictureAsLink(filename: string, insert?: boolean): number;
    addPictureAsLinkSync(filename: string, insert?: boolean): number;
    addPictureAsLinkAsync(filename: string, callback: (err: Error | null, id?: number) => void): Book;
    addPictureAsLinkAsync(filename: string, insert: boolean, callback: (err: Error | null, id?: number) => void): Book;

    // Default font
    defaultFont(): { name: string; size: number };
    setDefaultFont(name: string, size: number): Book;

    // Reference mode
    refR1C1(): boolean;
    setRefR1C1(refR1C1?: boolean): Book;

    // RGB mode
    rgbMode(): boolean;
    setRgbMode(rgbMode?: boolean): Book;

    // Version info
    version(): number;
    biffVersion(): number;

    // Date 1904 mode
    isDate1904(): boolean;
    setDate1904(date1904?: boolean): Book;

    // Template
    isTemplate(): boolean;
    setTemplate(isTemplate?: boolean): Book;

    // License key
    setKey(name: string, key: string): Book;

    // Write protection
    isWriteProtected(): boolean;

    // Core properties
    coreProperties(): CoreProperties;

    // Locale
    setLocale(locale: string): Book;

    // Cleanup
    removeVBA(): Book;
    removePrinterSettings(): Book;
    removeAllPhonetics(): Book;

    // DPI awareness
    dpiAwareness(): boolean;
    setDpiAwareness(dpiAwareness: boolean): Book;

    // Password
    setPassword(password: string): Book;

    // Error code
    errorCode(): number;

    // Clear
    clear(): Book;
}
