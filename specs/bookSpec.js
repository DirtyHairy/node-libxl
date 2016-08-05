var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    fs = require('fs'),
    shouldThrow = testUtils.shouldThrow;

testUtils.initFilesystem();

describe('The book class', function() {
    var book;

    beforeEach(function() {
        book = new xl.Book(xl.BOOK_TYPE_XLS);
    });

    it('Books are created via the book constructor', function() {
        expect(function() {var book = new xl.Book();}).toThrow();
        expect(function() {var book = new xl.Book(200);}).toThrow();
        expect(function() {var book = new xl.Book('a');}).toThrow();

        var bookXls = new xl.Book(xl.BOOK_TYPE_XLS),
            bookXlsx = new xl.Book(xl.BOOK_TYPE_XLSX);
    });

    it('book.writeSync writes a book in sync mode', function() {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        var file = testUtils.getWriteTestFile();
        shouldThrow(book.writeSync, book, 10);
        shouldThrow(book.writeSync, {}, file);
        expect(book.writeSync(file)).toBe(book);
    });

    it('book.write writes a book in sync mode', function() {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo'),
            result = null;

        sheet.writeStr(1, 0, 'bar');

        runs(function() {
            var file = testUtils.getWriteTestFile();
            shouldThrow(book.write, book, file, 10);
            shouldThrow(book.write, {}, file, function() {});

            expect(book.write(file, function(res) {
                result = res;
            })).toBe(book);

            shouldThrow(book.sheetCount, book);
        });

        waitsFor(function() {
            return result !== null;
        }, 'book to save', 1000);

        runs(function() {
            expect(result).toBeUndefined();
        });
    });

    it('book.loadSync loads a book in sync mode', function() {
        var file = testUtils.getWriteTestFile();
        shouldThrow(book.loadSync, book, 10);
        shouldThrow(book.loadSync, {}, file);

        expect(book.loadSync(file)).toBe(book);
        var sheet = book.getSheet(0);
        expect(sheet.readStr(1, 0)).toBe('bar');
    });

    it('book.load loads a book in async mode', function() {
        var result = null;

        runs(function() {
            var file = testUtils.getWriteTestFile();
            shouldThrow(book.load, book, file, 10);
            shouldThrow(book.load, {}, file, function() {});

            expect(book.load(file, function(res) {
                result = res;
            })).toBe(book);

            shouldThrow(book.sheetCount, book);
        });

        waitsFor(function() {
            return result !== null;
        }, 'book to load', 1000);

        runs(function() {
            expect(result).toBeUndefined();
            expect(book.getSheet(0).readStr(1, 0)).toBe('bar');
        });
    });

    it('book.loadRawSync and book.saveRawSync load and save a book into a buffer, sync mode', function() {
        var book1 = new xl.Book(xl.BOOK_TYPE_XLS);
        var sheet = book1.addSheet('foo');

        sheet.writeStr(1, 0, 'bar');

        shouldThrow(book1.writeRawSync, {});
        var buffer = book1.writeRawSync();

        var book2 = new xl.Book(xl.BOOK_TYPE_XLS);
        shouldThrow(book2.loadRawSync, book2, 1);
        shouldThrow(book2.loadRawSync, {}, buffer);

        expect(book2.loadRawSync(buffer)).toBe(book2);
        expect(book2.sheetCount()).toBe(1);
        expect(book2.getSheet(0).readStr(1, 0)).toBe('bar');
    });

    it('book.loadRaw and book.saveRaw load and save a book into a buffer, async mode', function() {
        var book;

        runs(function() {
            var book1 = new xl.Book(xl.BOOK_TYPE_XLS),
                book2 = new xl.Book(xl.BOOK_TYPE_XLS);
            var sheet = book1.addSheet('foo');

            sheet.writeStr(1, 0, 'bar');

            shouldThrow(book1.writeRaw, book1, 1);
            shouldThrow(book1.writeRaw, {}, function() {});

            function step1(err, buffer) {
                expect(err).toBeUndefined();

                shouldThrow(book2.loadRaw, book2, buffer, 10);
                shouldThrow(book2.loadRaw, {}, buffer, function() {});

                book2.loadRaw(buffer, step2);

                shouldThrow(book2.sheetCount, book2);
            }

            function step2(err) {
                expect(err).toBeUndefined();
                book = book2;
            }

            book1.writeRaw(step1);

            shouldThrow(book1.sheetCount, book1);
        });

        waitsFor(function() {
            return !!book;
        }, 'book to save and load again', 1000);

        runs(function() {
            expect(book.sheetCount()).toBe(1);
            expect(book.getSheet(0).readStr(1, 0)).toBe('bar');
        });
    });


    it('book.addSheet adds a sheet to a book', function() {
        shouldThrow(book.addSheet, book, 10);
        shouldThrow(book.addSheet, book, 'foo', 10);
        shouldThrow(book.addSheet, {}, 'foo');
        book.addSheet('baz');

        var sheet1 = book.addSheet('foo');
        sheet1.writeStr(1, 0, 'aaa');
        var sheet2 = book.addSheet('bar', sheet1);
        expect(sheet2.readStr(1, 0)).toBe('aaa');
    });

    it('book.insertSheet inserts a sheet at a given position', function() {
        var sheet1 = book.addSheet('foo');
        sheet1.writeStr(1, 0, 'bbb');

        shouldThrow(book.insertSheet, book, 'a', 'bar');
        shouldThrow(book.insertSheet, book, 2, 'bar');
        shouldThrow(book.insertSheet, {}, 0, 'bar');
        book.insertSheet(0, 'bar');
        expect(book.sheetCount()).toBe(2);

        var sheet2 = book.insertSheet(0, 'baz', sheet1);
        expect(sheet2.readStr(1, 0)).toBe('bbb');
    });

    it('book.getSheet retrieves a sheet at a given index', function() {
        var sheet = book.addSheet('foo');
        sheet.writeStr(1, 0, 'bar');

        shouldThrow(book.getSheet, book, 'a');
        shouldThrow(book.getSheet, book, 1);
        shouldThrow(book.getSheet, {}, 0);

        var sheet2 = book.getSheet(0);
        expect(sheet2.readStr(1, 0)).toBe('bar');
    });

    it('book.sheetType determines sheet type', function() {
        var sheet = book.addSheet('foo');
        shouldThrow(book.sheetType, book, 'a');
        shouldThrow(book.sheetType, {}, 0);
        expect(book.sheetType(0)).toBe(xl.SHEETTYPE_SHEET);
        expect(book.sheetType(1)).toBe(xl.SHEETTYPE_UNKNOWN);
    });

    it('book.delSheet removes a sheet', function() {
        book.addSheet('foo');
        book.addSheet('bar');
        shouldThrow(book.delSheet, book, 'a');
        shouldThrow(book.delSheet, book, 3);
        shouldThrow(book.delSheet, {}, 0);
        expect(book.delSheet(0)).toBe(book);
        expect(book.sheetCount()).toBe(1);
    });

    it('book.sheetCount counts the number of sheets in a book', function() {
        shouldThrow(book.sheetCount, {});
        expect(book.sheetCount()).toBe(0);
        book.addSheet('foo');
        expect(book.sheetCount()).toBe(1);
    });

    it('book.addFormat adds a format', function() {
        shouldThrow(book.addFormat, book, 10);
        shouldThrow(book.addFormat, {});
        book.addFormat();
        // TODO add a check for format inheritance once format has been
        // completed
    });

    it('book.addFormat can use a format belonging to a different book as a template', function() {
        var otherBook = new xl.Book(xl.BOOK_TYPE_XLS),
            format1 = otherBook.addFormat(),
            format2 = book.addFormat(format1);
    });

    it ('book.addFormat will throw if an async operation is pending on the template\'s parent book', function() {
        var otherBook = new xl.Book(xl.BOOK_TYPE_XLS),
            format1 = otherBook.addFormat();

        otherBook.saveRaw(function() {}, function() {});

        shouldThrow(book.addFormat, book, format1);
    });

    it('book.addFont adds a font', function() {
        shouldThrow(book.addFont, book, 10);
        shouldThrow(book.addFont, {});
        book.addFont();

        var font1 = book.addFont();
        font1.setName('times');
        expect(book.addFont(font1).name()).toBe('times');
    });

    it('book.addCusomtNumFormat adds a custom number format', function() {
        shouldThrow(book.addCusomtNumFormat, book, 10);
        shouldThrow(book.addCusomtNumFormat, {}, '000');
        book.addCustomNumFormat('000');
    });

    it('book.customNumFormat retrieves a custom number format', function() {
        var format = book.addCustomNumFormat('000');
        shouldThrow(book.customNumFormat, book, 'a');
        shouldThrow(book.customNumFormat, {}, format);

        expect(book.customNumFormat(format)).toBe('000');
    });

    it('book.format retrieves a format by index', function() {
        var format = book.addFormat();
        shouldThrow(book.format, book, 'a');
        shouldThrow(book.format, {}, 0);
        expect(book.format(0) instanceof format.constructor).toBe(true);
    });

    it('book.formatSize counts the number of formats', function() {
        shouldThrow(book.formatSize, {});
        var n = book.formatSize();
        book.addFormat();
        expect(book.formatSize()).toBe(n+1);
    });

    it('book.font gets a font by index', function() {
        var font = book.addFont();
        shouldThrow(book.font, book, -1);
        shouldThrow(book.font, {}, 0);
        expect(book.font(0) instanceof font.constructor).toBe(true);
    });

    it('book.fontSize counts the number of fonts', function() {
        shouldThrow(book.fontSize, {});
        var n = book.fontSize();
        book.addFont();
        expect(book.fontSize()).toBe(n+1);
    });

    it('book.datePack packs a date', function() {
        shouldThrow(book.datePack, book, 'a', 1, 2);
        shouldThrow(book.datePack, {}, 1, 2, 3);
        book.datePack(1, 2, 3);
    });

    it('book.datePack unpacks a date', function() {
        var date = book.datePack(1999, 2, 3, 4, 5, 6, 7);
        shouldThrow(book.dateUnpack, book, 'a');
        shouldThrow(book.dateUnpack, {}, date);

        var unpacked = book.dateUnpack(date);
        expect(unpacked.year).toBe(1999);
        expect(unpacked.month).toBe(2);
        expect(unpacked.day).toBe(3);
        expect(unpacked.hour).toBe(4);
        expect(unpacked.minute).toBe(5);
        expect(unpacked.second).toBe(6);
        expect(unpacked.msecond).toBe(7);
    });

    it('book.colorPack packs a color', function() {
        shouldThrow(book.colorPack, book, 'a', 2, 3);
        shouldThrow(book.colorPack, {}, 1, 2, 3);
        book.colorPack(1, 2, 3);
    });

    it('book.colorUnpack unpacks a color', function() {
        var color = book.colorPack(1, 2, 3);
        shouldThrow(book.colorUnpack, book, 'a');
        shouldThrow(book.colorUnpack, {}, color);

        var unpacked = book.colorUnpack(color);
        expect(unpacked.red).toBe(1);
        expect(unpacked.green).toBe(2);
        expect(unpacked.blue).toBe(3);
    });

    it('book.activeSheet returns the index of a book\'s active sheet', function() {
        book.addSheet('foo');
        book.addSheet('bar');
        book.setActiveSheet(0);
        shouldThrow(book.activeSheet, {});
        expect(book.activeSheet()).toBe(0);
        book.setActiveSheet(1);
        expect(book.activeSheet()).toBe(1);
    });

    it('book.setActiveSheet sets the active sheet', function() {
        book.addSheet('foo');
        shouldThrow(book.setActiveSheet, book, 'a');
        shouldThrow(book.setActiveSheet, {}, 0);
        expect(book.setActiveSheet(0)).toBe(book);
    });

    it('book.addPicture, book.pictureSize and book.getPicture ' +
            'manage the pictures in a book',
        function()
    {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            file = testUtils.getTestPicturePath(),
            fileBuffer = fs.readFileSync(file);

        shouldThrow(book.pictureSize, {});
        expect(book.pictureSize()).toBe(0);

        shouldThrow(book.addPicture, book, 1);
        shouldThrow(book.addPicture, {}, file);
        expect(book.addPicture(file)).toBe(0);
        expect(book.pictureSize()).toBe(1);

        shouldThrow(book.getPicture, book, true);
        shouldThrow(book.getPicture, {}, 0);
        var pic0 = book.getPicture(0);

        expect(pic0.type).toBe(xl.PICTURETYPE_PNG);
        expect(testUtils.compareBuffers(pic0.data, fileBuffer)).toBe(true);

        expect(book.addPicture(fileBuffer)).toBe(1);
        expect(book.pictureSize()).toBe(2);

        var pic1 = book.getPicture(1);
        expect(pic1.type).toBe(xl.PICTURETYPE_PNG);
        expect(testUtils.compareBuffers(pic1.data, fileBuffer)).toBe(true);
    });

    it('book.addPictureAsync and boook.addPictureAsync provide async picture management',
        function()
    {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            file = testUtils.getTestPicturePath(),
            fileBuffer = fs.readFileSync(file),
            buffer1, buffer2;

        runs(function() {
            function addFromFile(callback) {
                shouldThrow(book.addPictureAsync, book, file, 1);
                shouldThrow(book.addPictureAsync, {}, file, function() {});

                book.addPictureAsync(file, function(err, id) {
                    expect(err).toBeUndefined();
                    expect(id).toBe(0);
                    expect(book.pictureSize()).toBe(1);

                    addFromStream();
                });
            }

            function addFromStream() {
                book.addPictureAsync(fileBuffer, function(err, id) {
                    expect(err).toBeUndefined();
                    expect(id).toBe(1);
                    expect(book.pictureSize()).toBe(2);

                    getBuffer1();
                });
            }


            function getBuffer1() {
                shouldThrow(book.getPictureAsync, book, 1, 1);
                shouldThrow(book.getPictureAsync, {}, 1, function() {});

                book.getPictureAsync(0, function(err, type, data) {
                    expect(err).toBeUndefined();
                    expect(type).toBe(xl.PICTURETYPE_PNG);
                    buffer1 = data;

                    getBuffer2();
                });
            }

            function getBuffer2() {
                 book.getPictureAsync(1, function(err, type, data) {
                    expect(err).toBeUndefined();
                    expect(type).toBe(xl.PICTURETYPE_PNG);
                    buffer2 = data;
                });
            }

            addFromFile();
        });

        waitsFor(function() {
            return buffer1 && buffer2;
        }, 'pictures to load', 2000);

        runs(function() {
            expect(testUtils.compareBuffers(buffer1, fileBuffer)).toBe(true);
            expect(testUtils.compareBuffers(buffer2, fileBuffer)).toBe(true);
        });
    });

    it('book.defaultFont returns the default font', function() {
        book.setDefaultFont('times', 13);
        shouldThrow(book.defaultFont, {});

        var defaultFont = book.defaultFont();
        expect(defaultFont.name).toBe('times');
        expect(defaultFont.size).toBe(13);
    });

    it('book.setDefaultFont sets the default font', function() {
        shouldThrow(book.setDefaultFont, book, 10, 10);
        shouldThrow(book.setDefaultFont, {}, 'times', 10);
        expect(book.setDefaultFont('times', 10)).toBe(book);
    });

    it('book.refR1C1 checks for R1C1 mode', function() {
        book.setRefR1C1();
        shouldThrow(book.refR1C1, {});
        expect(book.refR1C1()).toBe(true);
        book.setRefR1C1(false);
        expect(book.refR1C1()).toBe(false);
    });

    it('book.setRefR1C1 switches R1C1 mode', function() {
        shouldThrow(book.setRefR1C1, book, 1);
        shouldThrow(book.setRefR1C1, {});
        expect(book.setRefR1C1(true)).toBe(book);
    });

    it('book.rgbMode checks for RGB mode', function() {
        book.setRgbMode();
        shouldThrow(book.rgbMode, {});
        expect(book.rgbMode()).toBe(true);
        book.setRgbMode(false);
        expect(book.rgbMode()).toBe(false);
    });

    it('book.setRgbMode switches RGB mode', function() {
        shouldThrow(book.setRgbMode, book, 1);
        shouldThrow(book.setRgbMode, {});
        expect(book.setRgbMode(true)).toBe(book);
    });

    it('book.biffVersion returns the BIFF format version', function() {
        shouldThrow(book.biffVersion, {});
        expect(book.biffVersion()).toBe(1536);
    });

    it('book.isDate1904 checks which date system is active', function() {
        book.setDate1904();
        shouldThrow(book.isDate1904, {});
        expect(book.isDate1904()).toBe(true);
        book.setDate1904(false);
        expect(book.isDate1904()).toBe(false);
    });

    it('book.setDate1904 sets the active date system', function() {
        shouldThrow(book.setDate1904, book, 1);
        shouldThrow(book.setDate1904, {}, true);
        expect(book.setDate1904(true)).toBe(book);
    });

    it('book.isTemplate checks whether the document is a template', function() {
        book.setTemplate();
        shouldThrow(book.isTemplate, {});
        expect(book.isTemplate()).toBe(true);
        book.setTemplate(false);
        expect(book.isTemplate()).toBe(false);
    });

    it('book.setTemplate toggles whether a document is a template', function() {
        shouldThrow(book.setTemplate, book, 1);
        shouldThrow(book.setTemplate, {});
        expect(book.setTemplate(true)).toBe(book);
    });

    it('book.setKey sets the API key', function() {
        shouldThrow(book.setKey, book, 1, 'a');
        shouldThrow(book.setKey, {}, 'a', 'b');
        expect(book.setKey('a', 'b')).toBe(book);
    });
});
