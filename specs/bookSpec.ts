import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import util from 'util';
import fs from 'fs';
import * as xl from '../lib/libxl';
import { initFilesystem, getWriteTestFile, getTempFile, getTestPicturePath, compareBuffers } from './testUtils';

initFilesystem();

describe('The book class', () => {
    let book: xl.Book;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLS);
    });

    it('Books are created via the book constructor', () => {
        assert.throws(() => {
            const book = new (xl.Book as any)();
        });
        assert.throws(() => {
            const book = new (xl.Book as any)(200);
        });
        assert.throws(() => {
            const book = new (xl.Book as any)('a');
        });
    });

    describe('book.writeSync', () => {
        it('writes a book in sync mode', () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            const file = getWriteTestFile();
            assert.throws(() => (book.writeSync as any).call(book, 10));
            assert.throws(() => (book.writeSync as any).call({}, file));
            assert.strictEqual(book.writeSync(file), book);
        });

        it('supports a tempfile', () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            const file = getWriteTestFile();
            assert.throws(() => (book.writeSync as any).call(book, 10));
            assert.throws(() => (book.writeSync as any).call({}, file));
            assert.strictEqual(book.writeSync(file), book, true as any);
        });
    });

    describe('book.write', () => {
        it('writes a book in async mode', async () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            const file = getWriteTestFile();
            assert.throws(() => (book.write as any).call(book, file, 10));
            assert.throws(() => (book.write as any).call({}, file, () => {}));

            const writeResult = util.promisify(book.write.bind(book))(file);
            assert.throws(() => (book.sheetCount as any).call(book));

            await writeResult;
        });

        it('supports a tempfile', async () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            const file = getWriteTestFile();
            assert.throws(() => (book.write as any).call(book, file, 10));
            assert.throws(() => (book.write as any).call({}, file, () => {}));
            assert.throws(() => (book.write as any).call(book, file, undefined, undefined, () => {}));

            const writeResult = util.promisify((cb) => book.write(file, true, cb))();
            assert.throws(() => (book.sheetCount as any).call(book));

            await writeResult;
        });
    });

    describe('book.loadSync', () => {
        it('loads a book in sync mode', () => {
            const file = getWriteTestFile();
            assert.throws(() => (book.loadSync as any).call(book, 10));
            assert.throws(() => (book.loadSync as any).call({}, file));

            assert.strictEqual(book.loadSync(file), book);
            const sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'bar');
        });

        it('supports an optional tempfile', () => {
            assert.strictEqual(book.loadSync(getWriteTestFile(), getTempFile()), book);
            const sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'bar');
        });
    });

    describe('book.load', () => {
        it('loads a book in async mode', async () => {
            const file = getWriteTestFile();
            assert.throws(() => (book.load as any).call(book, file, 10));
            assert.throws(() => (book.load as any).call({}, file, () => {}));
            assert.throws(() => (book.load as any).call(book, file, undefined, undefined, () => {}));

            const loadResult = util.promisify(book.load.bind(book))(file);
            assert.throws(() => (book.sheetCount as any).call(book));

            await loadResult;

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'bar');
        });

        it('supports an optional tempfile', async () => {
            await util.promisify((cb) => book.load(getWriteTestFile(), getTempFile(), cb))();

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'bar');
        });
    });

    describe('book.loadSheetSync', () => {
        it('loads a book in sync mode with a single sheet', () => {
            const file = getWriteTestFile();
            assert.throws(() => (book.loadSheetSync as any).call(book, file, 'a'));
            assert.throws(() => (book.loadSheetSync as any).call({}, file, 1));

            assert.strictEqual(book.loadSheetSync(file, 1), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports an optional tempfile', () => {
            const file = getWriteTestFile();

            assert.strictEqual(book.loadSheetSync(file, 1, getTempFile()), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', () => {
            const file = getWriteTestFile();

            assert.strictEqual(book.loadSheetSync(file, 1, undefined, true), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    describe('book.loadSheet', () => {
        it('loads a book in with a single sheet', async () => {
            const file = getWriteTestFile();
            assert.throws(() => (book.loadSheet as any).call(book, file, 'a', () => undefined));
            assert.throws(() => (book.loadSheet as any).call({}, file, 1, () => undefined));
            assert.throws(() =>
                (book.loadSheet as any).call(book, file, 1, undefined, undefined, undefined, () => undefined),
            );

            const loadResult = util.promisify(book.loadSheet.bind(book))(file, 1);
            assert.throws(() => (book.sheetCount as any).call(book));

            await loadResult;

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports an optional tempfile', async () => {
            const file = getWriteTestFile();

            await util.promisify((cb) => book.loadSheet(file, 1, getTempFile(), cb))();

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', async () => {
            const file = getWriteTestFile();

            await util.promisify((cb) => book.loadSheet(file, 1, cb))();

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    describe('book.loadPartiallySync', () => {
        it('loads a book in sync mode with a single, partial sheet', () => {
            const file = getWriteTestFile();
            assert.throws(() => (book.loadPartiallySync as any).call(book, file, 1, 0, 'a'));
            assert.throws(() => (book.loadPartiallySync as any).call({}, file, 1, 0, 1));

            assert.strictEqual(book.loadPartiallySync(file, 1, 0, 1), book);

            const sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'foo');
            assert.throws(() => (sheet.readStr as any).call(sheet, 2, 0));
        });

        it('supports an optional tempfile', () => {
            const file = getWriteTestFile();

            assert.strictEqual(book.loadPartiallySync(file, 1, 0, 1, getTempFile()), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', () => {
            const file = getWriteTestFile();

            assert.strictEqual(book.loadPartiallySync(file, 1, 0, 1, undefined, true), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    describe('book.loadPartially', () => {
        it('loads a book with a single, partial sheet', async () => {
            const file = getWriteTestFile();
            assert.throws(() => (book.loadPartially as any).call(book, file, 1, 0, 'a', () => undefined));
            assert.throws(() => (book.loadPartially as any).call({}, file, 1, 0, 1, () => undefined));
            assert.throws(() =>
                (book.loadPartially as any).call(book, file, 1, 0, 1, undefined, undefined, undefined, () => undefined),
            );

            const loadResult = util.promisify((cb) => book.loadPartially(file, 1, 0, 1, cb))();
            assert.throws(() => (book.sheetCount as any).call(book));

            await loadResult;

            const sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'foo');
            assert.throws(() => (sheet.readStr as any).call(sheet, 2, 0));
        });

        it('supports an optional tempfile', async () => {
            const file = getWriteTestFile();

            await util.promisify((cb) => book.loadPartially(file, 1, 0, 1, getTempFile(), cb))();

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', async () => {
            const file = getWriteTestFile();

            await util.promisify((cb) => book.loadPartially(file, 1, 0, 1, getTempFile(), true, cb))();

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    it('book.loadWithoutEmptyCellsSync loads a book in sync mode without empty cells', () => {
        const file = getWriteTestFile();
        assert.throws(() => (book.loadWithoutEmptyCellsSync as any).call(book, 10));
        assert.throws(() => (book.loadWithoutEmptyCellsSync as any).call({}, file));

        assert.strictEqual(book.loadWithoutEmptyCellsSync(file), book);
        const sheet = book.getSheet(0);
        assert.strictEqual(sheet.readStr(1, 0), 'bar');
    });

    it('book.loadWithoutEmptyCells loads a book in async mode without empty cells', async () => {
        const file = getWriteTestFile();
        assert.throws(() => (book.loadWithoutEmptyCells as any).call(book, file, 10));
        assert.throws(() => (book.loadWithoutEmptyCells as any).call({}, file, () => {}));
        assert.throws(() => (book.loadWithoutEmptyCells as any).call(book, file, undefined, () => {}));

        const loadResult = util.promisify(book.loadWithoutEmptyCells.bind(book))(file);
        assert.throws(() => (book.sheetCount as any).call(book));

        await loadResult;

        assert.strictEqual(book.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadInfoSync loads a book in sync mode without data', () => {
        const file = getWriteTestFile();
        assert.throws(() => (book.loadInfoSync as any).call(book, 10));
        assert.throws(() => (book.loadInfoSync as any).call({}, file));

        assert.strictEqual(book.loadInfoSync(file), book);
        assert.strictEqual(book.sheetCount(), 2);
    });

    it('book.loadInfo loads a book in async mode without data', async () => {
        const file = getWriteTestFile();
        assert.throws(() => (book.loadInfo as any).call(book, file, 10));
        assert.throws(() => (book.loadInfo as any).call({}, file, () => {}));
        assert.throws(() => (book.loadInfo as any).call(book, file, undefined, () => {}));

        const loadResult = util.promisify(book.loadInfo.bind(book))(file);
        assert.throws(() => (book.sheetCount as any).call(book));

        await loadResult;

        assert.strictEqual(book.sheetCount(), 2);
    });

    it('book.loadRawSync and book.saveRawSync load and save a book into a buffer, sync mode', () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        const sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        assert.throws(() => (book1.writeRawSync as any).call({}));
        const buffer = book1.writeRawSync();

        const book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        assert.throws(() => (book2.loadRawSync as any).call(book2, 1));
        assert.throws(() => (book2.loadRawSync as any).call({}, buffer));

        assert.strictEqual(book2.loadRawSync(buffer), book2);
        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadRawSync supports extra args for sheet and range', () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        const sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        const buffer = book1.writeRawSync();

        const book2 = new xl.Book(xl.BOOK_TYPE_XLS);

        assert.strictEqual(book2.loadRawSync(buffer, 0, 0, 1, true), book2);
        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadRaw and book.saveRaw load and save a book into a buffer, async mode', async () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS),
            book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        const sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        assert.throws(() => (book1.writeRaw as any).call(book1, 1));
        assert.throws(() => (book1.writeRaw as any).call({}, () => {}));

        const writeResult = util.promisify((cb: (err: Error | null, buffer: Buffer) => void) => {
            book1.writeRaw(cb);
        })();
        assert.throws(() => (book1.sheetCount as any).call(book1));

        const buffer = await writeResult;

        assert.throws(() => (book2.loadRaw as any).call(book2, buffer, 10));
        assert.throws(() => (book2.loadRaw as any).call({}, buffer, () => {}));
        assert.throws(() =>
            (book2.loadRaw as any).call(book2, buffer, undefined, undefined, undefined, undefined, undefined, () => {}),
        );

        const loadResult = util.promisify((cb) => {
            book2.loadRaw(buffer, cb);
        })();
        assert.throws(() => (book2.loadRaw as any).call(book2));
        await loadResult;

        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadRaw supports extra args for sheet and range', async () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS),
            book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        const sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        const buffer = await util.promisify((cb: (err: Error | null, buffer: Buffer) => void) => {
            book1.writeRaw(cb);
        })();

        await util.promisify((cb) => {
            book2.loadRaw(buffer, 0, 0, 1, true, cb);
        })();

        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.addSheet adds a sheet to a book', () => {
        assert.throws(() => (book.addSheet as any).call(book, 10));
        assert.throws(() => (book.addSheet as any).call(book, 'foo', 10));
        assert.throws(() => (book.addSheet as any).call({}, 'foo'));
        book.addSheet('baz');

        const sheet1 = book.addSheet('foo');
        sheet1.writeStr(1, 0, 'aaa');
        const sheet2 = book.addSheet('bar', sheet1);
        assert.strictEqual(sheet2.readStr(1, 0), 'aaa');
    });

    it('book.insertSheet inserts a sheet at a given position', () => {
        const sheet1 = book.addSheet('foo');
        sheet1.writeStr(1, 0, 'bbb');

        assert.throws(() => (book.insertSheet as any).call(book, 'a', 'bar'));
        assert.throws(() => (book.insertSheet as any).call(book, 2, 'bar'));
        assert.throws(() => (book.insertSheet as any).call({}, 0, 'bar'));
        book.insertSheet(0, 'bar');
        assert.strictEqual(book.sheetCount(), 2);

        const sheet2 = book.insertSheet(0, 'baz', sheet1);
        assert.strictEqual(sheet2.readStr(1, 0), 'bbb');
    });

    it('book.getSheet retrieves a sheet at a given index', () => {
        const sheet = book.addSheet('foo');
        sheet.writeStr(1, 0, 'bar');

        assert.throws(() => (book.getSheet as any).call(book, 'a'));
        assert.throws(() => (book.getSheet as any).call(book, 1));
        assert.throws(() => (book.getSheet as any).call({}, 0));

        const sheet2 = book.getSheet(0);
        assert.strictEqual(sheet2.readStr(1, 0), 'bar');
    });

    it('book.getSheetName retrieves the name of a sheet', () => {
        book.addSheet('foo');

        assert.throws(() => (book.getSheetName as any).call(book, 'a'));
        assert.throws(() => (book.getSheetName as any).call(book, 1));
        assert.throws(() => (book.getSheetName as any).call({}, 0));

        assert.strictEqual(book.getSheetName(0), 'foo');
    });

    it('book.sheetType determines sheet type', () => {
        const sheet = book.addSheet('foo');
        assert.throws(() => (book.sheetType as any).call(book, 'a'));
        assert.throws(() => (book.sheetType as any).call({}, 0));
        assert.strictEqual(book.sheetType(0), xl.SHEETTYPE_SHEET);
        assert.strictEqual(book.sheetType(1), xl.SHEETTYPE_UNKNOWN);
    });

    it('book.moveSheet moves a sheet', () => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);

        book.addSheet('foo');
        book.addSheet('bar');
        book.addSheet('baz');

        assert.throws(() => (book.moveSheet as any).call(book, 'a', 0));
        assert.throws(() => (book.moveSheet as any).call({}, 1, 0));

        assert.strictEqual(book.moveSheet(2, 0), book);
        assert.strictEqual(book.getSheetName(0), 'baz');
    });

    it('book.delSheet removes a sheet', () => {
        book.addSheet('foo');
        book.addSheet('bar');
        assert.throws(() => (book.delSheet as any).call(book, 'a'));
        assert.throws(() => (book.delSheet as any).call(book, 3));
        assert.throws(() => (book.delSheet as any).call({}, 0));
        assert.strictEqual(book.delSheet(0), book);
        assert.strictEqual(book.sheetCount(), 1);
    });

    it('book.sheetCount counts the number of sheets in a book', () => {
        assert.throws(() => (book.sheetCount as any).call({}));
        assert.strictEqual(book.sheetCount(), 0);
        book.addSheet('foo');
        assert.strictEqual(book.sheetCount(), 1);
    });

    it('book.addFormat adds a format', () => {
        assert.throws(() => (book.addFormat as any).call(book, 10));
        assert.throws(() => (book.addFormat as any).call({}));

        book.addFormat();
    });

    it('book.addFormat can use a format belonging to a different book as a template', () => {
        const otherBook = new xl.Book(xl.BOOK_TYPE_XLS),
            format1 = otherBook.addFormat(),
            format2 = book.addFormat(format1);
    });

    it("book.addFormat will throw if an async operation is pending on the template's parent book", () => {
        const otherBook = new xl.Book(xl.BOOK_TYPE_XLS),
            format1 = otherBook.addFormat();

        (otherBook as any).saveRaw(
            () => {},
            () => {},
        );

        assert.throws(() => (book.addFormat as any).call(book, format1));
    });

    it('book.addFormatFromStyle adds a format from a cell style', () => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);

        assert.throws(() => (book.addFormatFromStyle as any).call(book, 'a'));
        assert.throws(() => (book.addFormatFromStyle as any).call({}, xl.CELLSTYLE_NORMAL));

        assert.strictEqual(book.addFormatFromStyle(xl.CELLSTYLE_BAD) instanceof book.addFormat().constructor, true);
    });

    it('book.addFont adds a font', () => {
        assert.throws(() => (book.addFont as any).call(book, 10));
        assert.throws(() => (book.addFont as any).call({}));
        book.addFont();

        const font1 = book.addFont();
        font1.setName('times');
        assert.strictEqual(book.addFont(font1).name(), 'times');
    });

    it('book.addConditionalFormat adds a conditional format', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        assert.throws(() => (book.addConditionalFormat as any).call({}));

        assert.ok(book.addConditionalFormat() instanceof xl.ConditionalFormat);
    });

    it('book.addRichString adds a rich string', () => {
        assert.throws(() => (book.addRichString as any).call({}));

        const richString = book.addRichString();
        assert.strictEqual(richString instanceof xl.RichString, true);
    });

    it('book.addCustomNumFormat adds a custom number format', () => {
        assert.throws(() => (book.addCustomNumFormat as any).call(book, 10));
        assert.throws(() => (book.addCustomNumFormat as any).call({}, '000'));
        book.addCustomNumFormat('000');
    });

    it('book.customNumFormat retrieves a custom number format', () => {
        const format = book.addCustomNumFormat('000');
        assert.throws(() => (book.customNumFormat as any).call(book, 'a'));
        assert.throws(() => (book.customNumFormat as any).call({}, format));

        assert.strictEqual(book.customNumFormat(format), '000');
    });

    it('book.format retrieves a format by index', () => {
        const format = book.addFormat();
        assert.throws(() => (book.format as any).call(book, 'a'));
        assert.throws(() => (book.format as any).call({}, 0));
        assert.strictEqual(book.format(0) instanceof xl.Format, true);
    });

    it('book.formatSize counts the number of formats', () => {
        assert.throws(() => (book.formatSize as any).call({}));
        const n = book.formatSize();
        book.addFormat();
        assert.strictEqual(book.formatSize(), n + 1);
    });

    it('book.font gets a font by index', () => {
        const font = book.addFont();
        assert.throws(() => (book.font as any).call(book, -1));
        assert.throws(() => (book.font as any).call({}, 0));
        assert.strictEqual(book.font(0) instanceof xl.Font, true);
    });

    it('book.fontSize counts the number of fonts', () => {
        assert.throws(() => (book.fontSize as any).call({}));
        const n = book.fontSize();
        book.addFont();
        assert.strictEqual(book.fontSize(), n + 1);
    });

    it('book.datePack packs a date', () => {
        assert.throws(() => (book.datePack as any).call(book, 'a', 1, 2));
        assert.throws(() => (book.datePack as any).call({}, 1, 2, 3));
        book.datePack(1, 2, 3);
    });

    it('book.datePack unpacks a date', () => {
        const date = book.datePack(1999, 2, 3, 4, 5, 6, 7);
        assert.throws(() => (book.dateUnpack as any).call(book, 'a'));
        assert.throws(() => (book.dateUnpack as any).call({}, date));

        const unpacked = book.dateUnpack(date);
        assert.strictEqual(unpacked.year, 1999);
        assert.strictEqual(unpacked.month, 2);
        assert.strictEqual(unpacked.day, 3);
        assert.strictEqual(unpacked.hour, 4);
        assert.strictEqual(unpacked.minute, 5);
        assert.strictEqual(unpacked.second, 6);
        assert.strictEqual(unpacked.msecond, 7);
    });

    it('book.colorPack packs a color', () => {
        assert.throws(() => (book.colorPack as any).call(book, 'a', 2, 3));
        assert.throws(() => (book.colorPack as any).call({}, 1, 2, 3));
        book.colorPack(1, 2, 3);
    });

    it('book.colorUnpack unpacks a color', () => {
        const color = book.colorPack(1, 2, 3);
        assert.throws(() => (book.colorUnpack as any).call(book, 'a'));
        assert.throws(() => (book.colorUnpack as any).call({}, color));

        const unpacked = book.colorUnpack(color);
        assert.strictEqual(unpacked.red, 1);
        assert.strictEqual(unpacked.green, 2);
        assert.strictEqual(unpacked.blue, 3);
    });

    it("book.activeSheet returns the index of a book's active sheet", () => {
        book.addSheet('foo');
        book.addSheet('bar');
        book.setActiveSheet(0);
        assert.throws(() => (book.activeSheet as any).call({}));
        assert.strictEqual(book.activeSheet(), 0);
        book.setActiveSheet(1);
        assert.strictEqual(book.activeSheet(), 1);
    });

    it('book.setActiveSheet sets the active sheet', () => {
        book.addSheet('foo');
        assert.throws(() => (book.setActiveSheet as any).call(book, 'a'));
        assert.throws(() => (book.setActiveSheet as any).call({}, 0));
        assert.strictEqual(book.setActiveSheet(0), book);
    });

    it('book.addPicture, book.pictureSize and book.getPicture ' + 'manage the pictures in a book', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLS),
            file = getTestPicturePath(),
            fileBuffer = fs.readFileSync(file);

        assert.throws(() => (book.pictureSize as any).call({}));
        assert.strictEqual(book.pictureSize(), 0);

        assert.throws(() => (book.addPicture as any).call(book, 1));
        assert.throws(() => (book.addPicture as any).call({}, file));
        assert.strictEqual(book.addPicture(file), 0);
        assert.strictEqual(book.pictureSize(), 1);

        assert.throws(() => (book.getPicture as any).call(book, true));
        assert.throws(() => (book.getPicture as any).call({}, 0));
        const pic0 = book.getPicture(0);

        assert.strictEqual(pic0.type, xl.PICTURETYPE_JPEG);
        assert.strictEqual(compareBuffers(pic0.data, fileBuffer), true);

        assert.strictEqual(book.addPicture(fileBuffer), 1);
        assert.strictEqual(book.pictureSize(), 2);

        const pic1 = book.getPicture(1);
        assert.strictEqual(pic1.type, xl.PICTURETYPE_JPEG);
        assert.strictEqual(compareBuffers(pic1.data, fileBuffer), true);
    });

    it('book.addPictureAsync and book.getPictureAsync provide async picture management', async () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLS),
            file = getTestPicturePath(),
            fileBuffer = fs.readFileSync(file);

        assert.throws(() => (book.addPictureAsync as any).call(book, file, 1));
        assert.throws(() => (book.addPictureAsync as any).call({}, file, () => {}));

        const addResult = util.promisify((cb) => book.addPictureAsync(file, cb))();
        assert.throws(() => (book.sheetCount as any).call(book));

        assert.strictEqual(await addResult, 0);
        assert.strictEqual(book.pictureSize(), 1);

        assert.strictEqual(await util.promisify((cb) => book.addPictureAsync(fileBuffer, cb))(), 1);
        assert.strictEqual(book.pictureSize(), 2);

        const getPictureAsync = (id: number) =>
            new Promise<[number, Buffer]>((resolve, reject) =>
                book.getPictureAsync(id, (err: any, type: number, data: Buffer) =>
                    err ? reject(err) : resolve([type, data]),
                ),
            );

        const getResult = getPictureAsync(0);
        assert.throws(() => (book.sheetCount as any).call(book));

        const [type1, buffer1] = await getResult;
        const [type2, buffer2] = await getPictureAsync(1);

        assert.strictEqual(type1, xl.PICTURETYPE_JPEG);
        assert.strictEqual(type2, xl.PICTURETYPE_JPEG);

        assert.strictEqual(compareBuffers(buffer1, fileBuffer), true);
        assert.strictEqual(compareBuffers(buffer2, fileBuffer), true);
    });

    describe('book.addPictureAsLinkSync', () => {
        it('adds a picture as link', () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = getTestPicturePath();

            assert.throws(() => (book.addPictureAsLinkSync as any).call(book, 1));
            assert.throws(() => (book.addPictureAsLinkSync as any).call({}, file));

            assert.strictEqual(book.addPictureAsLink(file), 0);
            assert.strictEqual(book.pictureSize(), 1);
        });

        it('accepts a second parameter insert', () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = getTestPicturePath();

            assert.strictEqual(book.addPictureAsLink(file, true), 0);
            assert.strictEqual(book.pictureSize(), 1);
        });
    });

    describe('book.addPictureAsLinkAsync', () => {
        it('adds a picture as link', async () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = getTestPicturePath();

            assert.throws(() => (book.addPictureAsLinkAsync as any).call(book, 1, () => undefined));
            assert.throws(() => (book.addPictureAsLinkAsync as any).call({}, file, () => undefined));

            const result = util.promisify((cb) => book.addPictureAsLinkAsync(file, cb))();
            assert.throws(() => (book.sheetCount as any).call(book));

            assert.strictEqual(await result, 0);
            assert.strictEqual(book.pictureSize(), 1);
        });

        it('supports a second parameter insert', async () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = getTestPicturePath();

            const result = util.promisify((cb) => book.addPictureAsLinkAsync(file, true, cb))();
            assert.throws(() => (book.sheetCount as any).call(book));

            assert.strictEqual(await result, 0);
            assert.strictEqual(book.pictureSize(), 1);
        });
    });

    it('book.defaultFont returns the default font', () => {
        book.setDefaultFont('times', 13);
        assert.throws(() => (book.defaultFont as any).call({}));

        const defaultFont = book.defaultFont();
        assert.strictEqual(defaultFont.name, 'times');
        assert.strictEqual(defaultFont.size, 13);
    });

    it('book.setDefaultFont sets the default font', () => {
        assert.throws(() => (book.setDefaultFont as any).call(book, 10, 10));
        assert.throws(() => (book.setDefaultFont as any).call({}, 'times', 10));
        assert.strictEqual(book.setDefaultFont('times', 10), book);
    });

    it('book.refR1C1 checks for R1C1 mode', () => {
        book.setRefR1C1();
        assert.throws(() => (book.refR1C1 as any).call({}));
        assert.strictEqual(book.refR1C1(), true);
        book.setRefR1C1(false);
        assert.strictEqual(book.refR1C1(), false);
    });

    it('book.setRefR1C1 switches R1C1 mode', () => {
        assert.throws(() => (book.setRefR1C1 as any).call(book, 1));
        assert.throws(() => (book.setRefR1C1 as any).call({}));
        assert.strictEqual(book.setRefR1C1(true), book);
    });

    it('book.rgbMode checks for RGB mode', () => {
        book.setRgbMode();
        assert.throws(() => (book.rgbMode as any).call({}));
        assert.strictEqual(book.rgbMode(), true);
        book.setRgbMode(false);
        assert.strictEqual(book.rgbMode(), false);
    });

    it('book.setRgbMode switches RGB mode', () => {
        assert.throws(() => (book.setRgbMode as any).call(book, 1));
        assert.throws(() => (book.setRgbMode as any).call({}));
        assert.strictEqual(book.setRgbMode(true), book);
    });

    it('book.version returns the libxl version', () => {
        assert.throws(() => (book.version as any).call({}));
        assert.strictEqual(typeof book.version(), 'number');
    });

    it('book.biffVersion returns the BIFF format version', () => {
        assert.throws(() => (book.biffVersion as any).call({}));
        assert.strictEqual(book.biffVersion(), 1536);
    });

    it('book.isDate1904 checks which date system is active', () => {
        book.setDate1904();
        assert.throws(() => (book.isDate1904 as any).call({}));
        assert.strictEqual(book.isDate1904(), true);
        book.setDate1904(false);
        assert.strictEqual(book.isDate1904(), false);
    });

    it('book.setDate1904 sets the active date system', () => {
        assert.throws(() => (book.setDate1904 as any).call(book, 1));
        assert.throws(() => (book.setDate1904 as any).call({}, true));
        assert.strictEqual(book.setDate1904(true), book);
    });

    it('book.isTemplate checks whether the document is a template', () => {
        book.setTemplate();
        assert.throws(() => (book.isTemplate as any).call({}));
        assert.strictEqual(book.isTemplate(), true);
        book.setTemplate(false);
        assert.strictEqual(book.isTemplate(), false);
    });

    it('book.setTemplate toggles whether a document is a template', () => {
        assert.throws(() => (book.setTemplate as any).call(book, 1));
        assert.throws(() => (book.setTemplate as any).call({}));
        assert.strictEqual(book.setTemplate(true), book);
    });

    it('book.setKey sets the API key', () => {
        assert.throws(() => (book.setKey as any).call(book, 1, 'a'));
        assert.throws(() => (book.setKey as any).call({}, 'a', 'b'));
        assert.strictEqual(book.setKey('a', 'b'), book);
    });

    it('book.isWriteProtected checks for write protection', () => {
        assert.throws(() => (book.isWriteProtected as any).call({}));
        assert.strictEqual(book.isWriteProtected(), false);
    });

    it('book.getCoreProperty retrieves core properties', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        assert.throws(() => (book.coreProperties as any).call({}));

        assert.ok(book.coreProperties() instanceof xl.CoreProperties);
    });

    it('book.setLocale sets the locale', () => {
        const locale = process.env['LANG'] ?? 'en_GB';

        assert.throws(() => (book.setLocale as any).call(book, 1));
        assert.throws(() => (book.setLocale as any).call({}, locale));

        assert.strictEqual(book.setLocale(locale), book);
    });

    it('book.remveVBA removes all VBA macros', () => {
        assert.throws(() => (book.removeVBA as any).call({}));

        assert.strictEqual(book.removeVBA(), book);
    });

    it('book.removePrinterSettings removes printer settings', () => {
        assert.throws(() => (book.removePrinterSettings as any).call({}));

        assert.strictEqual(book.removePrinterSettings(), book);
    });

    it('book.removeAllPhonetics removes all phonetics', () => {
        assert.throws(() => (book.removeAllPhonetics as any).call({}));

        assert.strictEqual(book.removeAllPhonetics(), book);
    });

    it('book.dpiAwareness and book.setDpiAwareness mange DPI awareness', () => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);

        assert.throws(() => (book.dpiAwareness as any).call({}));
        assert.throws(() => (book.setDpiAwareness as any).call(book, 1));
        assert.throws(() => (book.setDpiAwareness as any).call({}, false));

        assert.strictEqual(book.setDpiAwareness(false), book);
        assert.strictEqual(book.dpiAwareness(), false);
        assert.strictEqual(book.setDpiAwareness(true), book);
        assert.strictEqual(book.dpiAwareness(), true);
    });

    it('book.setPassword sets a password for encrypted XLSX files', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        assert.throws(() => (book.setPassword as any).call(book, 1));
        assert.throws(() => (book.setPassword as any).call({}, 'secret'));

        assert.strictEqual(book.setPassword('secret'), book);
    });

    it('book.loadInfoRawSync loads info from a buffer in sync mode', () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        book1.addSheet('testSheet');

        const buffer = book1.writeRawSync();

        const book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        assert.throws(() => (book2.loadInfoRawSync as any).call(book2, 1));
        assert.throws(() => (book2.loadInfoRawSync as any).call({}, buffer));

        assert.strictEqual(book2.loadInfoRawSync(buffer), book2);
        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheetName(0), 'testSheet');
    });

    it('book.loadInfoRaw loads info from a buffer in async mode', async () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        book1.addSheet('asyncSheet');

        const buffer = book1.writeRawSync();

        const book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        assert.throws(() => (book2.loadInfoRaw as any).call(book2, buffer, 10));
        assert.throws(() => (book2.loadInfoRaw as any).call({}, buffer, () => {}));

        const loadResult = util.promisify(book2.loadInfoRaw.bind(book2))(buffer);
        assert.throws(() => (book2.loadInfoRaw as any).call(book2));
        await loadResult;

        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheetName(0), 'asyncSheet');
    });

    it('book.errorCode returns an error code number', () => {
        assert.throws(() => (book.errorCode as any).call({}));
        assert.strictEqual(typeof book.errorCode(), 'number');
        assert.strictEqual(book.errorCode(), xl.ERRCODE_OK);
    });

    it('book.conditionalFormatSize and book.conditionalFormat manage conditional formats', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        assert.throws(() => (book.conditionalFormatSize as any).call({}));
        assert.strictEqual(book.conditionalFormatSize(), 0);

        const cf = book.addConditionalFormat();
        assert.strictEqual(book.conditionalFormatSize(), 1);

        assert.throws(() => (book.conditionalFormat as any).call(book, 'a'));
        assert.throws(() => (book.conditionalFormat as any).call({}, 0));

        assert.ok(book.conditionalFormat(0) instanceof xl.ConditionalFormat);
    });

    it('book.clear clears all workbook data', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.addSheet('sheet1');
        book.addSheet('sheet2');

        assert.throws(() => (book.clear as any).call({}));

        assert.strictEqual(book.clear(), book);
        assert.strictEqual(book.sheetCount(), 0);
    });
});
