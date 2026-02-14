const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    fs = require('fs'),
    shouldThrow = testUtils.shouldThrow;

testUtils.initFilesystem();

describe('The book class', function () {
    var book;

    beforeEach(function () {
        book = new xl.Book(xl.BOOK_TYPE_XLS);
    });

    it('Books are created via the book constructor', function () {
        assert.throws(function () {
            var book = new xl.Book();
        });
        assert.throws(function () {
            var book = new xl.Book(200);
        });
        assert.throws(function () {
            var book = new xl.Book('a');
        });
    });

    describe('book.writeSync', () => {
        it('writes a book in sync mode', function () {
            var book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            var file = testUtils.getWriteTestFile();
            shouldThrow(book.writeSync, book, 10);
            shouldThrow(book.writeSync, {}, file);
            assert.strictEqual(book.writeSync(file), book);
        });

        it('supports a tempfile', function () {
            var book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            var file = testUtils.getWriteTestFile();
            shouldThrow(book.writeSync, book, 10);
            shouldThrow(book.writeSync, {}, file);
            assert.strictEqual(book.writeSync(file), book, true);
        });
    });

    describe('book.write', () => {
        it('writes a book in async mode', async () => {
            var book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            var file = testUtils.getWriteTestFile();
            shouldThrow(book.write, book, file, 10);
            shouldThrow(book.write, {}, file, function () {});

            const writeResult = util.promisify(book.write.bind(book))(file);
            shouldThrow(book.sheetCount, book);

            await writeResult;
        });

        it('supports a tempfile', async () => {
            var book = new xl.Book(xl.BOOK_TYPE_XLS),
                sheet1 = book.addSheet('foo'),
                sheet2 = book.addSheet('bar');

            sheet1.writeStr(1, 0, 'bar');
            sheet2.writeStr(1, 0, 'foo');
            sheet2.writeStr(2, 0, 'row2');

            var file = testUtils.getWriteTestFile();
            shouldThrow(book.write, book, file, 10);
            shouldThrow(book.write, {}, file, function () {});
            shouldThrow(book.write, book, file, undefined, undefined, function () {});

            const writeResult = util.promisify(book.write.bind(book))(file, true);
            shouldThrow(book.sheetCount, book);

            await writeResult;
        });
    });

    describe('book.loadSync', () => {
        it('loads a book in sync mode', function () {
            var file = testUtils.getWriteTestFile();
            shouldThrow(book.loadSync, book, 10);
            shouldThrow(book.loadSync, {}, file);

            assert.strictEqual(book.loadSync(file), book);
            var sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'bar');
        });

        it('supports an optional tempfile', function () {
            assert.strictEqual(book.loadSync(testUtils.getWriteTestFile(), testUtils.getTempFile()), book);
            var sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'bar');
        });
    });

    describe('book.load', () => {
        it('loads a book in async mode', async () => {
            var file = testUtils.getWriteTestFile();
            shouldThrow(book.load, book, file, 10);
            shouldThrow(book.load, {}, file, function () {});
            shouldThrow(book.load, book, file, undefined, undefined, function () {});

            const loadResult = util.promisify(book.load.bind(book))(file);
            shouldThrow(book.sheetCount, book);

            await loadResult;

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'bar');
        });

        it('supports an optional tempfile', async () => {
            await util.promisify(book.load.bind(book))(testUtils.getWriteTestFile(), testUtils.getTempFile());

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'bar');
        });
    });

    describe('book.loadSheetSync', () => {
        it('loads a book in sync mode with a single sheet', () => {
            const file = testUtils.getWriteTestFile();
            shouldThrow(book.loadSyncSheet, book, file, 'a');
            shouldThrow(book.loadSyncSheet, {}, file, 1);

            assert.strictEqual(book.loadSheetSync(file, 1), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports an optional tempfile', () => {
            const file = testUtils.getWriteTestFile();

            assert.strictEqual(book.loadSheetSync(file, 1, testUtils.getTempFile()), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', () => {
            const file = testUtils.getWriteTestFile();

            assert.strictEqual(book.loadSheetSync(file, 1, undefined, true), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    describe('book.loadSheet', () => {
        it('loads a book in with a single sheet', async () => {
            const file = testUtils.getWriteTestFile();
            shouldThrow(book.loadSheet, book, file, 'a', () => undefined);
            shouldThrow(book.loadSheet, {}, file, 1, () => undefined);
            shouldThrow(book.loadSheet, book, file, 1, undefined, undefined, undefined, () => undefined);

            const loadResult = util.promisify(book.loadSheet.bind(book))(file, 1);
            shouldThrow(book.sheetCount, book);

            await loadResult;

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports an optional tempfile', async () => {
            const file = testUtils.getWriteTestFile();

            await util.promisify(book.loadSheet.bind(book))(file, 1, testUtils.getTempFile());

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', async () => {
            const file = testUtils.getWriteTestFile();

            await util.promisify(book.loadSheet.bind(book))(file, 1, undefined, true);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    describe('book.loadPartiallySync', () => {
        it('loads a book in sync mode with a single, partial sheet', () => {
            const file = testUtils.getWriteTestFile();
            shouldThrow(book.loadPartiallySync, book, file, 1, 0, 'a');
            shouldThrow(book.loadPartiallySync, {}, file, 1, 0, 1);

            assert.strictEqual(book.loadPartiallySync(file, 1, 0, 1), book);

            const sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'foo');
            shouldThrow(sheet.readStr, sheet, 2, 0);
        });

        it('supports an optional tempfile', () => {
            const file = testUtils.getWriteTestFile();

            assert.strictEqual(book.loadPartiallySync(file, 1, 0, 1, testUtils.getTempFile()), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', () => {
            const file = testUtils.getWriteTestFile();

            assert.strictEqual(book.loadPartiallySync(file, 1, 0, 1, undefined, true), book);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    describe('book.loadPartially', () => {
        it('loads a book with a single, partial sheet', async () => {
            const file = testUtils.getWriteTestFile();
            shouldThrow(book.loadPartially, book, file, 1, 0, 'a', () => undefined);
            shouldThrow(book.loadPartially, {}, file, 1, 0, 1, () => undefined);
            shouldThrow(book.loadPartially, book, file, 1, 0, 1, undefined, undefined, undefined, () => undefined);

            const loadResult = util.promisify(book.loadPartially.bind(book))(file, 1, 0, 1);
            shouldThrow(book.sheetCount, book);

            await loadResult;

            const sheet = book.getSheet(0);
            assert.strictEqual(sheet.readStr(1, 0), 'foo');
            shouldThrow(sheet.readStr, sheet, 2, 0);
        });

        it('supports an optional tempfile', async () => {
            const file = testUtils.getWriteTestFile();

            await util.promisify(book.loadPartially.bind(book))(file, 1, 0, 1, testUtils.getTempFile());

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });

        it('supports loading all sheets', async () => {
            const file = testUtils.getWriteTestFile();

            await util.promisify(book.loadPartially.bind(book))(file, 1, 0, 1, undefined, true);

            assert.strictEqual(book.getSheet(0).readStr(1, 0), 'foo');
        });
    });

    it('book.loadWithoutEmptyCellsSync loads a book in sync mode without empty cells', function () {
        var file = testUtils.getWriteTestFile();
        shouldThrow(book.loadWithoutEmptyCellsSync, book, 10);
        shouldThrow(book.loadWithoutEmptyCellsSync, {}, file);

        assert.strictEqual(book.loadWithoutEmptyCellsSync(file), book);
        var sheet = book.getSheet(0);
        assert.strictEqual(sheet.readStr(1, 0), 'bar');
    });

    it('book.loadWithoutEmptyCells loads a book in async mode without empty cells', async () => {
        var file = testUtils.getWriteTestFile();
        shouldThrow(book.loadWithoutEmptyCells, book, file, 10);
        shouldThrow(book.loadWithoutEmptyCells, {}, file, function () {});
        shouldThrow(book.loadWithoutEmptyCells, book, file, undefined, function () {});

        const loadResult = util.promisify(book.loadWithoutEmptyCells.bind(book))(file);
        shouldThrow(book.sheetCount, book);

        await loadResult;

        assert.strictEqual(book.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadInfoSync loads a book in sync mode without data', function () {
        var file = testUtils.getWriteTestFile();
        shouldThrow(book.loadInfoSync, book, 10);
        shouldThrow(book.loadInfoSync, {}, file);

        assert.strictEqual(book.loadInfoSync(file), book);
        assert.strictEqual(book.sheetCount(), 2);
    });

    it('book.loadInfo loads a book in async mode without data', async () => {
        var file = testUtils.getWriteTestFile();
        shouldThrow(book.loadInfo, book, file, 10);
        shouldThrow(book.loadInfo, {}, file, function () {});
        shouldThrow(book.loadInfo, book, file, undefined, function () {});

        const loadResult = util.promisify(book.loadInfo.bind(book))(file);
        shouldThrow(book.sheetCount, book);

        await loadResult;

        assert.strictEqual(book.sheetCount(), 2);
    });

    it('book.loadRawSync and book.saveRawSync load and save a book into a buffer, sync mode', function () {
        var book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        var sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        shouldThrow(book1.writeRawSync, {});
        var buffer = book1.writeRawSync();

        var book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        shouldThrow(book2.loadRawSync, book2, 1);
        shouldThrow(book2.loadRawSync, {}, buffer);

        assert.strictEqual(book2.loadRawSync(buffer), book2);
        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadRawSync supports extra args for sheet and range', function () {
        var book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        var sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        var buffer = book1.writeRawSync();

        var book2 = new xl.Book(xl.BOOK_TYPE_XLS);

        assert.strictEqual(book2.loadRawSync(buffer, 0, 0, 1, true), book2);
        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadRaw and book.saveRaw load and save a book into a buffer, async mode', async () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS),
            book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        const sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        shouldThrow(book1.writeRaw, book1, 1);
        shouldThrow(book1.writeRaw, {}, function () {});

        const writeResult = util.promisify(book1.writeRaw.bind(book1))();
        shouldThrow(book1.sheetCount, book1);

        const buffer = await writeResult;

        shouldThrow(book2.loadRaw, book2, buffer, 10);
        shouldThrow(book2.loadRaw, {}, buffer, function () {});
        shouldThrow(
            book2.loadRaw,
            book2,
            buffer,
            undefined,
            undefined,
            undefined,
            undefined,
            undefined,
            function () {},
        );

        const loadResult = util.promisify(book2.loadRaw.bind(book2))(buffer);
        shouldThrow(book2.loadRaw, book2);
        await loadResult;

        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.loadRaw supports extra args for sheet and range', async () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS),
            book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        const sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        const buffer = await util.promisify(book1.writeRaw.bind(book1))();

        await util.promisify(book2.loadRaw.bind(book2))(buffer, 0, 0, 1, true);

        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheet(0).readStr(1, 0), 'bar');
    });

    it('book.addSheet adds a sheet to a book', function () {
        shouldThrow(book.addSheet, book, 10);
        shouldThrow(book.addSheet, book, 'foo', 10);
        shouldThrow(book.addSheet, {}, 'foo');
        book.addSheet('baz');

        var sheet1 = book.addSheet('foo');
        sheet1.writeStr(1, 0, 'aaa');
        var sheet2 = book.addSheet('bar', sheet1);
        assert.strictEqual(sheet2.readStr(1, 0), 'aaa');
    });

    it('book.insertSheet inserts a sheet at a given position', function () {
        var sheet1 = book.addSheet('foo');
        sheet1.writeStr(1, 0, 'bbb');

        shouldThrow(book.insertSheet, book, 'a', 'bar');
        shouldThrow(book.insertSheet, book, 2, 'bar');
        shouldThrow(book.insertSheet, {}, 0, 'bar');
        book.insertSheet(0, 'bar');
        assert.strictEqual(book.sheetCount(), 2);

        var sheet2 = book.insertSheet(0, 'baz', sheet1);
        assert.strictEqual(sheet2.readStr(1, 0), 'bbb');
    });

    it('book.getSheet retrieves a sheet at a given index', function () {
        var sheet = book.addSheet('foo');
        sheet.writeStr(1, 0, 'bar');

        shouldThrow(book.getSheet, book, 'a');
        shouldThrow(book.getSheet, book, 1);
        shouldThrow(book.getSheet, {}, 0);

        var sheet2 = book.getSheet(0);
        assert.strictEqual(sheet2.readStr(1, 0), 'bar');
    });

    it('book.getSheetName retrieves the name of a sheet', function () {
        book.addSheet('foo');

        shouldThrow(book.getSheetName, book, 'a');
        shouldThrow(book.getSheetName, book, 1);
        shouldThrow(book.getSheetName, {}, 0);

        assert.strictEqual(book.getSheetName(0), 'foo');
    });

    it('book.sheetType determines sheet type', function () {
        var sheet = book.addSheet('foo');
        shouldThrow(book.sheetType, book, 'a');
        shouldThrow(book.sheetType, {}, 0);
        assert.strictEqual(book.sheetType(0), xl.SHEETTYPE_SHEET);
        assert.strictEqual(book.sheetType(1), xl.SHEETTYPE_UNKNOWN);
    });

    it('book.moveSheet moves a sheet', function () {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);

        book.addSheet('foo');
        book.addSheet('bar');
        book.addSheet('baz');

        shouldThrow(book.moveSheet, book, 'a', 0);
        shouldThrow(book.moveSheet, {}, 1, 0);

        assert.strictEqual(book.moveSheet(2, 0), book);
        assert.strictEqual(book.getSheetName(0), 'baz');
    });

    it('book.delSheet removes a sheet', function () {
        book.addSheet('foo');
        book.addSheet('bar');
        shouldThrow(book.delSheet, book, 'a');
        shouldThrow(book.delSheet, book, 3);
        shouldThrow(book.delSheet, {}, 0);
        assert.strictEqual(book.delSheet(0), book);
        assert.strictEqual(book.sheetCount(), 1);
    });

    it('book.sheetCount counts the number of sheets in a book', function () {
        shouldThrow(book.sheetCount, {});
        assert.strictEqual(book.sheetCount(), 0);
        book.addSheet('foo');
        assert.strictEqual(book.sheetCount(), 1);
    });

    it('book.addFormat adds a format', function () {
        shouldThrow(book.addFormat, book, 10);
        shouldThrow(book.addFormat, {});

        book.addFormat();
    });

    it('book.addFormat can use a format belonging to a different book as a template', function () {
        var otherBook = new xl.Book(xl.BOOK_TYPE_XLS),
            format1 = otherBook.addFormat(),
            format2 = book.addFormat(format1);
    });

    it("book.addFormat will throw if an async operation is pending on the template's parent book", function () {
        var otherBook = new xl.Book(xl.BOOK_TYPE_XLS),
            format1 = otherBook.addFormat();

        otherBook.saveRaw(
            function () {},
            function () {},
        );

        shouldThrow(book.addFormat, book, format1);
    });

    it('book.addFormatFromStyle adds a format from a cell style', function () {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);

        shouldThrow(book.addFormatFromStyle, book, 'a');
        shouldThrow(book.addFormatFromStyle, {}, xl.CELLSTYLE_NORMAL);

        assert.strictEqual(book.addFormatFromStyle(xl.CELLSTYLE_BAD) instanceof book.addFormat().constructor, true);
    });

    it('book.addFont adds a font', function () {
        shouldThrow(book.addFont, book, 10);
        shouldThrow(book.addFont, {});
        book.addFont();

        var font1 = book.addFont();
        font1.setName('times');
        assert.strictEqual(book.addFont(font1).name(), 'times');
    });

    it('book.addConditionalFormat adds a conditional format', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        shouldThrow(book.addConditionalFormat, {});

        assert.ok(book.addConditionalFormat() instanceof xl.ConditionalFormat);
    });

    it('book.addRichString adds a rich string', () => {
        shouldThrow(book.addRichString, {});

        const richString = book.addRichString();
        assert.strictEqual(richString instanceof xl.RichString, true);
    });

    it('book.addCusomtNumFormat adds a custom number format', function () {
        shouldThrow(book.addCusomtNumFormat, book, 10);
        shouldThrow(book.addCusomtNumFormat, {}, '000');
        book.addCustomNumFormat('000');
    });

    it('book.customNumFormat retrieves a custom number format', function () {
        var format = book.addCustomNumFormat('000');
        shouldThrow(book.customNumFormat, book, 'a');
        shouldThrow(book.customNumFormat, {}, format);

        assert.strictEqual(book.customNumFormat(format), '000');
    });

    it('book.format retrieves a format by index', function () {
        var format = book.addFormat();
        shouldThrow(book.format, book, 'a');
        shouldThrow(book.format, {}, 0);
        assert.strictEqual(book.format(0) instanceof xl.Format, true);
    });

    it('book.formatSize counts the number of formats', function () {
        shouldThrow(book.formatSize, {});
        var n = book.formatSize();
        book.addFormat();
        assert.strictEqual(book.formatSize(), n + 1);
    });

    it('book.font gets a font by index', function () {
        var font = book.addFont();
        shouldThrow(book.font, book, -1);
        shouldThrow(book.font, {}, 0);
        assert.strictEqual(book.font(0) instanceof xl.Font, true);
    });

    it('book.fontSize counts the number of fonts', function () {
        shouldThrow(book.fontSize, {});
        var n = book.fontSize();
        book.addFont();
        assert.strictEqual(book.fontSize(), n + 1);
    });

    it('book.datePack packs a date', function () {
        shouldThrow(book.datePack, book, 'a', 1, 2);
        shouldThrow(book.datePack, {}, 1, 2, 3);
        book.datePack(1, 2, 3);
    });

    it('book.datePack unpacks a date', function () {
        var date = book.datePack(1999, 2, 3, 4, 5, 6, 7);
        shouldThrow(book.dateUnpack, book, 'a');
        shouldThrow(book.dateUnpack, {}, date);

        var unpacked = book.dateUnpack(date);
        assert.strictEqual(unpacked.year, 1999);
        assert.strictEqual(unpacked.month, 2);
        assert.strictEqual(unpacked.day, 3);
        assert.strictEqual(unpacked.hour, 4);
        assert.strictEqual(unpacked.minute, 5);
        assert.strictEqual(unpacked.second, 6);
        assert.strictEqual(unpacked.msecond, 7);
    });

    it('book.colorPack packs a color', function () {
        shouldThrow(book.colorPack, book, 'a', 2, 3);
        shouldThrow(book.colorPack, {}, 1, 2, 3);
        book.colorPack(1, 2, 3);
    });

    it('book.colorUnpack unpacks a color', function () {
        var color = book.colorPack(1, 2, 3);
        shouldThrow(book.colorUnpack, book, 'a');
        shouldThrow(book.colorUnpack, {}, color);

        var unpacked = book.colorUnpack(color);
        assert.strictEqual(unpacked.red, 1);
        assert.strictEqual(unpacked.green, 2);
        assert.strictEqual(unpacked.blue, 3);
    });

    it("book.activeSheet returns the index of a book's active sheet", function () {
        book.addSheet('foo');
        book.addSheet('bar');
        book.setActiveSheet(0);
        shouldThrow(book.activeSheet, {});
        assert.strictEqual(book.activeSheet(), 0);
        book.setActiveSheet(1);
        assert.strictEqual(book.activeSheet(), 1);
    });

    it('book.setActiveSheet sets the active sheet', function () {
        book.addSheet('foo');
        shouldThrow(book.setActiveSheet, book, 'a');
        shouldThrow(book.setActiveSheet, {}, 0);
        assert.strictEqual(book.setActiveSheet(0), book);
    });

    it('book.addPicture, book.pictureSize and book.getPicture ' + 'manage the pictures in a book', function () {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            file = testUtils.getTestPicturePath(),
            fileBuffer = fs.readFileSync(file);

        shouldThrow(book.pictureSize, {});
        assert.strictEqual(book.pictureSize(), 0);

        shouldThrow(book.addPicture, book, 1);
        shouldThrow(book.addPicture, {}, file);
        assert.strictEqual(book.addPicture(file), 0);
        assert.strictEqual(book.pictureSize(), 1);

        shouldThrow(book.getPicture, book, true);
        shouldThrow(book.getPicture, {}, 0);
        var pic0 = book.getPicture(0);

        assert.strictEqual(pic0.type, xl.PICTURETYPE_JPEG);
        assert.strictEqual(testUtils.compareBuffers(pic0.data, fileBuffer), true);

        assert.strictEqual(book.addPicture(fileBuffer), 1);
        assert.strictEqual(book.pictureSize(), 2);

        var pic1 = book.getPicture(1);
        assert.strictEqual(pic1.type, xl.PICTURETYPE_JPEG);
        assert.strictEqual(testUtils.compareBuffers(pic1.data, fileBuffer), true);
    });

    it('book.addPictureAsync and book.getPictureAsync provide async picture management', async () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLS),
            file = testUtils.getTestPicturePath(),
            fileBuffer = fs.readFileSync(file);

        shouldThrow(book.addPictureAsync, book, file, 1);
        shouldThrow(book.addPictureAsync, {}, file, function () {});

        const addPictureAsync = util.promisify(book.addPictureAsync.bind(book));

        const addResult = addPictureAsync(file);
        shouldThrow(book.sheetCount, book);

        assert.strictEqual(await addResult, 0);
        assert.strictEqual(book.pictureSize(), 1);

        assert.strictEqual(await addPictureAsync(fileBuffer), 1);
        assert.strictEqual(book.pictureSize(), 2);

        const getPictureAsync = (id) =>
            new Promise((resolve, reject) =>
                book.getPictureAsync(id, (err, type, data) => (err ? reject(err) : resolve([type, data]))),
            );

        const getResult = getPictureAsync(0);
        shouldThrow(book.sheetCount, book);

        const [type1, buffer1] = await getResult;
        const [type2, buffer2] = await getPictureAsync(1);

        assert.strictEqual(type1, xl.PICTURETYPE_JPEG);
        assert.strictEqual(type2, xl.PICTURETYPE_JPEG);

        assert.strictEqual(testUtils.compareBuffers(buffer1, fileBuffer), true);
        assert.strictEqual(testUtils.compareBuffers(buffer2, fileBuffer), true);
    });

    describe('book.addPictureAsLinkSync', () => {
        it('adds a picture as link', () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = testUtils.getTestPicturePath();

            shouldThrow(book.addPictureAsLinkSync, book, 1);
            shouldThrow(book.addPictureAsLinkSync, {}, file);

            assert.strictEqual(book.addPictureAsLink(file), 0);
            assert.strictEqual(book.pictureSize(), 1);
        });

        it('accepts a second parameter insert', () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = testUtils.getTestPicturePath();

            assert.strictEqual(book.addPictureAsLink(file, true), 0);
            assert.strictEqual(book.pictureSize(), 1);
        });
    });

    describe('book.addPictureAsLinkAsync', () => {
        it('adds a picture as link', async () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = testUtils.getTestPicturePath();

            shouldThrow(book.addPictureAsLinkAsync, book, 1, () => undefined);
            shouldThrow(book.addPictureAsLinkAsync, {}, file, () => undefined);

            const result = util.promisify(book.addPictureAsLinkAsync.bind(book))(file);
            shouldThrow(book.sheetCount, book);

            assert.strictEqual(await result, 0);
            assert.strictEqual(book.pictureSize(), 1);
        });

        it('supports a second parameter insert', async () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLSX);
            const file = testUtils.getTestPicturePath();

            const result = util.promisify(book.addPictureAsLinkAsync.bind(book))(file, true);
            shouldThrow(book.sheetCount, book);

            assert.strictEqual(await result, 0);
            assert.strictEqual(book.pictureSize(), 1);
        });
    });

    it('book.defaultFont returns the default font', function () {
        book.setDefaultFont('times', 13);
        shouldThrow(book.defaultFont, {});

        var defaultFont = book.defaultFont();
        assert.strictEqual(defaultFont.name, 'times');
        assert.strictEqual(defaultFont.size, 13);
    });

    it('book.setDefaultFont sets the default font', function () {
        shouldThrow(book.setDefaultFont, book, 10, 10);
        shouldThrow(book.setDefaultFont, {}, 'times', 10);
        assert.strictEqual(book.setDefaultFont('times', 10), book);
    });

    it('book.refR1C1 checks for R1C1 mode', function () {
        book.setRefR1C1();
        shouldThrow(book.refR1C1, {});
        assert.strictEqual(book.refR1C1(), true);
        book.setRefR1C1(false);
        assert.strictEqual(book.refR1C1(), false);
    });

    it('book.setRefR1C1 switches R1C1 mode', function () {
        shouldThrow(book.setRefR1C1, book, 1);
        shouldThrow(book.setRefR1C1, {});
        assert.strictEqual(book.setRefR1C1(true), book);
    });

    it('book.rgbMode checks for RGB mode', function () {
        book.setRgbMode();
        shouldThrow(book.rgbMode, {});
        assert.strictEqual(book.rgbMode(), true);
        book.setRgbMode(false);
        assert.strictEqual(book.rgbMode(), false);
    });

    it('book.setRgbMode switches RGB mode', function () {
        shouldThrow(book.setRgbMode, book, 1);
        shouldThrow(book.setRgbMode, {});
        assert.strictEqual(book.setRgbMode(true), book);
    });

    it('book.version returns the libxl version', function () {
        shouldThrow(book.version, {});
        assert.strictEqual(typeof book.version(), 'number');
    });

    it('book.biffVersion returns the BIFF format version', function () {
        shouldThrow(book.biffVersion, {});
        assert.strictEqual(book.biffVersion(), 1536);
    });

    it('book.isDate1904 checks which date system is active', function () {
        book.setDate1904();
        shouldThrow(book.isDate1904, {});
        assert.strictEqual(book.isDate1904(), true);
        book.setDate1904(false);
        assert.strictEqual(book.isDate1904(), false);
    });

    it('book.setDate1904 sets the active date system', function () {
        shouldThrow(book.setDate1904, book, 1);
        shouldThrow(book.setDate1904, {}, true);
        assert.strictEqual(book.setDate1904(true), book);
    });

    it('book.isTemplate checks whether the document is a template', function () {
        book.setTemplate();
        shouldThrow(book.isTemplate, {});
        assert.strictEqual(book.isTemplate(), true);
        book.setTemplate(false);
        assert.strictEqual(book.isTemplate(), false);
    });

    it('book.setTemplate toggles whether a document is a template', function () {
        shouldThrow(book.setTemplate, book, 1);
        shouldThrow(book.setTemplate, {});
        assert.strictEqual(book.setTemplate(true), book);
    });

    it('book.setKey sets the API key', function () {
        shouldThrow(book.setKey, book, 1, 'a');
        shouldThrow(book.setKey, {}, 'a', 'b');
        assert.strictEqual(book.setKey('a', 'b'), book);
    });

    it('book.isWriteProtected checks for write protection', () => {
        shouldThrow(book.isWriteProtected, {});
        assert.strictEqual(book.isWriteProtected(), false);
    });

    it('book.getCoreProperty retrieves core properties', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        shouldThrow(book.coreProperties, {});

        assert.ok(book.coreProperties() instanceof xl.CoreProperties);
    });

    it('book.setLocale sets the locale', () => {
        const locale = process.env['LANG'] ?? 'en_GB';

        shouldThrow(book.setLocale, book, 1);
        shouldThrow(book.setLocale, {}, locale);

        assert.strictEqual(book.setLocale(locale), book);
    });

    it('book.remveVBA removes all VBA macros', () => {
        shouldThrow(book.removeVBA, {});

        assert.strictEqual(book.removeVBA(), book);
    });

    it('book.removePrinterSettings removes printer settings', () => {
        shouldThrow(book.removePrinterSettings, {});

        assert.strictEqual(book.removePrinterSettings(), book);
    });

    it('book.removeAllPhonetics removes all phonetics', () => {
        shouldThrow(book.removeAllPhonetics, {});

        assert.strictEqual(book.removeAllPhonetics(), book);
    });

    it('book.dpiAwareness and book.setDpiAwareness mange DPI awareness', () => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);

        shouldThrow(book.dpiAwareness, {});
        shouldThrow(book.setDpiAwareness, book, 1);
        shouldThrow(book.setDpiAwareness, {}, false);

        assert.strictEqual(book.setDpiAwareness(false), book);
        assert.strictEqual(book.dpiAwareness(), false);
        assert.strictEqual(book.setDpiAwareness(true), book);
        assert.strictEqual(book.dpiAwareness(), true);
    });

    it('book.setPassword sets a password for encrypted XLSX files', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        shouldThrow(book.setPassword, book, 1);
        shouldThrow(book.setPassword, {}, 'secret');

        assert.strictEqual(book.setPassword('secret'), book);
    });

    it('book.loadInfoRawSync loads info from a buffer in sync mode', () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        book1.addSheet('testSheet');

        const buffer = book1.writeRawSync();

        const book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        shouldThrow(book2.loadInfoRawSync, book2, 1);
        shouldThrow(book2.loadInfoRawSync, {}, buffer);

        assert.strictEqual(book2.loadInfoRawSync(buffer), book2);
        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheetName(0), 'testSheet');
    });

    it('book.loadInfoRaw loads info from a buffer in async mode', async () => {
        const book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        book1.addSheet('asyncSheet');

        const buffer = book1.writeRawSync();

        const book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        shouldThrow(book2.loadInfoRaw, book2, buffer, 10);
        shouldThrow(book2.loadInfoRaw, {}, buffer, function () {});

        const loadResult = util.promisify(book2.loadInfoRaw.bind(book2))(buffer);
        shouldThrow(book2.loadInfoRaw, book2);
        await loadResult;

        assert.strictEqual(book2.sheetCount(), 1);
        assert.strictEqual(book2.getSheetName(0), 'asyncSheet');
    });

    it('book.errorCode returns an error code number', () => {
        shouldThrow(book.errorCode, {});
        assert.strictEqual(typeof book.errorCode(), 'number');
        assert.strictEqual(book.errorCode(), xl.ERRCODE_OK);
    });

    it('book.conditionalFormatSize and book.conditionalFormat manage conditional formats', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);

        shouldThrow(book.conditionalFormatSize, {});
        assert.strictEqual(book.conditionalFormatSize(), 0);

        const cf = book.addConditionalFormat();
        assert.strictEqual(book.conditionalFormatSize(), 1);

        shouldThrow(book.conditionalFormat, book, 'a');
        shouldThrow(book.conditionalFormat, {}, 0);

        assert.ok(book.conditionalFormat(0) instanceof xl.ConditionalFormat);
    });

    it('book.clear clears all workbook data', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.addSheet('sheet1');
        book.addSheet('sheet2');

        shouldThrow(book.clear, {});

        assert.strictEqual(book.clear(), book);
        assert.strictEqual(book.sheetCount(), 0);
    });
});
