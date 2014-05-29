var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils');

function shouldThrow(fun, scope) {
    var args = Array.prototype.slice.call(arguments, 2);

    expect(function() {fun.apply(scope, args);}).toThrow();
}

testUtils.initFilesystem();

describe('The font class', function() {

    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        font = book.addFont();

    it('The font constructor can not be called directly', function() {
        expect(function() {new font.constructor();}).toThrow();
    });

    it('font.size reads font size', function() {
        shouldThrow(font.size, book);
        font.setSize(10);
        expect(font.size()).toBe(10);
        font.setSize(20);
        expect(font.size()).toBe(20);
    });

    it('font.setSize sets font size', function() {
        shouldThrow(font.setSize, font, '');
        shouldThrow(font.setSize, {}, 10);
        expect(font.setSize(10)).toBe(font);
    });

    it('font.italic checks wether a font is italic', function() {
        shouldThrow(font.italic, book);
        font.setItalic(false);
        expect(font.italic()).toBe(false);
        font.setItalic(true);
        expect(font.italic()).toBe(true);
    });

    it('font.setItalic toggles a font italic', function() {
        shouldThrow(font.setItalic, font, 10);
        shouldThrow(font.setItalic, book, true);
        expect(font.setItalic(true)).toBe(font);
    });

    it('font.strikeOut checks wether a font is striked out', function() {
        shouldThrow(font.strikeOut, book);
        font.setStrikeOut(false);
        expect(font.strikeOut()).toBe(false);
        font.setStrikeOut(true);
        expect(font.strikeOut()).toBe(true);
    });

    it('font.setStrikeOut toggles a font striked out', function() {
        shouldThrow(font.setStrikeOut, font, 10);
        shouldThrow(font.setStrikeOut, book, true);
        expect(font.setStrikeOut(true)).toBe(font);
    });
   

    it('font.color gets font color', function() {
        shouldThrow(font.color, book);
        font.setColor(xl.COLOR_BLACK);
        expect(font.color()).toBe(xl.COLOR_BLACK);
        font.setColor(xl.COLOR_WHITE);
        expect(font.color()).toBe(xl.COLOR_WHITE);
    });

    it('font.setColor sets font color', function() {
        shouldThrow(font.setColor, font, true);
        shouldThrow(font.setColor, book, xl.COLOR_BLACK);
        expect(font.setColor(xl.COLOR_BLACK)).toBe(font);
    });

    it('font.bold checks whether a font is bold', function() {
        shouldThrow(font.bold, book);
        font.setBold(false);
        expect(font.bold()).toBe(false);
        font.setBold(true);
        expect(font.bold()).toBe(true);
    });

    it('font.setBold toggles a font bold', function() {
        shouldThrow(font.setBold, font, 10);
        shouldThrow(font.setBold, book, true);
        expect(font.setBold(true)).toBe(font);
    });

    it('font.script gets a fonts script style', function() {
        shouldThrow(font.script, book);
        font.setScript(xl.SCRIPT_NORMAL);
        expect(font.script()).toBe(xl.SCRIPT_NORMAL);
        font.setScript(xl.SCRIPT_SUPER);
        expect(font.script()).toBe(xl.SCRIPT_SUPER);
    });

    it('font.setScript sets a fonts script style', function() {
        shouldThrow(font.setScript, font, '');
        shouldThrow(font.setScript, book, xl.SCRIPT_NORMAL);
        expect(font.setScript(xl.SCRIPT_NORMAL)).toBe(font);
    });

    it('font.underline gets a fonts underline mode', function() {
        shouldThrow(font.underline, book);
        font.setUnderline(xl.UNDERLINE_NONE);
        expect(font.underline()).toBe(xl.UNDERLINE_NONE);
        font.setUnderline(xl.UNDERLINE_SINGLE);
        expect(font.underline()).toBe(xl.UNDERLINE_SINGLE);
    });

    it('font.setUnderline sets a fonts underline mode', function() {
        shouldThrow(font.setUnderline, font, true);
        shouldThrow(font.setUnderline, book, xl.UNDERLINE_NONE);
        expect(font.setUnderline(xl.UNDERLINE_NONE)).toBe(font);
    });

    it('font.name gets the font name', function() {
        shouldThrow(font.name, book);
        font.setName('arial');
        expect(font.name()).toBe('arial');
        font.setName('times');
        expect(font.name()).toBe('times');
    });

    it('font.setName sets the font name', function() {
        shouldThrow(font.setName, font, 10);
        shouldThrow(font.setName, book, 'arial');
        expect(font.setName('arial')).toBe(font);
    });
});

describe('The sheet class', function() {

    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        sheet = book.addSheet('foo'),
        format = book.addFormat(),
        wrongBook = new xl.Book(xl.BOOK_TYPE_XLS),
        wrongSheet = wrongBook.addSheet('bar'),
        wrongFormat = wrongBook.addFormat(),
        row = 1;

    it('sheet.cellType determines cell type', function() {
        sheet
            .writeStr(row, 0, 'foo')
            .writeNum(row, 1, 10);

        expect(function() {sheet.cellType();}).toThrow();
        expect(function() {sheet.cellType.call({}, row, 0);}).toThrow();

        expect(sheet.cellType(row, 0)).toBe(xl.CELLTYPE_STRING);
        expect(sheet.cellType(row, 1)).toBe(xl.CELLTYPE_NUMBER);

        row++;
    });

    it('sheet.isFormula checks whether a cell contains a formula', function() {
        sheet
            .writeStr(row, 0, 'foo')
            .writeFormula(row, 1, '=1');

        expect(function() {sheet.isFormula();}).toThrow();
        expect(function() {sheet.isFormula.call({}, row, 0);}).toThrow();

        expect(sheet.isFormula(row, 0)).toBe(false);
        expect(sheet.isFormula(row, 1)).toBe(true);

        row++;
    });

    it ('sheet.cellFormat retrieves the cell format', function() {
        expect(function() {sheet.cellFormat();}).toThrow();
        expect(function() {sheet.cellFormat.call({}, row, 0);}).toThrow();

        var cellFormat = sheet.cellFormat(row, 0);
        expect(format instanceof format.constructor).toBe(true);
    });

    it('sheet.setCellFormat sets the cell format', function() {
        expect(function() {sheet.setCellFormat();}).toThrow();
        expect(function() {sheet.setCellFormat.call({}, row, 0, format);}).toThrow();

        expect(function() {sheet.setCellFormat(row, 0, wrongFormat);}).toThrow();
        expect(sheet.setCellFormat(row, 0, format)).toBe(sheet);
    });

    it('sheet.readStr reads a string', function() {
        sheet.writeStr(row, 0, 'foo');
        sheet.writeNum(row, 1, 10);

        expect(function() {sheet.readStr();}).toThrow();
        expect(function() {sheet.readStr.call({}, row, 0);}).toThrow();

        var formatRef = {};
        expect(function() {sheet.readStr(row, 1);}).toThrow();
        expect(sheet.readStr(row, 0)).toBe('foo');
        expect(sheet.readStr(row, 0, formatRef)).toBe('foo');
        expect(formatRef.format instanceof format.constructor).toBe(true);

        row++;
    });

    it('sheet.writeStr writes a string', function() {
        expect(function() {sheet.writeStr();}).toThrow();
        expect(function() {sheet.writeStr.call({}, row, 0, 'foo');}).toThrow();

        expect(function() {sheet.writeStr(row, 0, 'foo', wrongFormat);}).toThrow();
        expect(sheet.writeStr(row, 0, 'foo')).toBe(sheet);
        expect(sheet.writeStr(row, 0, 'foo', format)).toBe(sheet);

        row++;
    });

    it('sheet.readNum reads a number', function() {
        sheet.writeNum(row, 0, 10);

        expect(function() {sheet.readNum();}).toThrow();
        expect(function() {sheet.readNum.call({}, row, 0);}).toThrow();

        var formatRef = {};
        expect(sheet.readNum(row, 0)).toEqual(10);
        expect(sheet.readNum(row, 0, formatRef)).toEqual(10);
        expect(formatRef.format instanceof format.constructor).toBe(true);

        row++;
    });

    it('sheet.writeNum writes a Number', function() {
        expect(function() {sheet.writeNum();}).toThrow();
        expect(function() {sheet.writeNum.call({}, row, 0, 10);}).toThrow();

        expect(function() {sheet.writeNum(row, 0, 10, wrongFormat);}).toThrow();
        expect(sheet.writeNum(row, 0, 10)).toBe(sheet);
        expect(sheet.writeNum(row, 0, 10, format)).toBe(sheet);

        row++;
    });

    it('sheet.readBool reads a bool', function() {
        sheet.writeBool(row, 0, true);

        expect(function() {sheet.readBool();}).toThrow();
        expect(function() {sheet.readBool.call({}, row, 0);}).toThrow();

        var formatRef = {};
        expect(sheet.readBool(row, 0)).toEqual(true);
        expect(sheet.readBool(row, 0, formatRef)).toEqual(true);
        expect(formatRef.format instanceof format.constructor).toBe(true);

        row++;
    });

    it('sheet.writeBool writes a bool', function() {
        expect(function() {sheet.writeBool();}).toThrow();
        expect(function() {sheet.writeBool.call({}, row, 0, true);}).toThrow();

        expect(function() {sheet.writeBool(row, 0, true, wrongFormat);}).toThrow();
        expect(sheet.writeBool(row, 0, true)).toBe(sheet);
        expect(sheet.writeBool(row, 0, true, format)).toBe(sheet);

        row++;
    });

    it('sheet.readBlank reads a blank cell', function() {
        sheet.writeBlank(row, 0, format);
        sheet.writeNum(row, 1, 10);

        expect(function() {sheet.readBlank();}).toThrow();
        expect(function() {sheet.readBlank.call({}, row, 0);}).toThrow();

        var formatRef = {};
        expect(sheet.readBlank(row, 0) instanceof format.constructor).toBe(true);
        expect(sheet.readBlank(row, 0, formatRef) instanceof format.constructor).toBe(true);
        expect(formatRef.format instanceof format.constructor).toBe(true);
        expect(function() {sheet.readBlank(row, 1);}).toThrow();

        row++;
    });

    it('sheet.writeBlank writes a blank cell', function() {
        expect(function() {sheet.writeBlank();}).toThrow();
        expect(function() {sheet.writeBlank.call({}, row, 0, format);}).toThrow();

        expect(function() {sheet.writeBlank(row, 0, wrongFormat);}).toThrow();
        expect(function() {sheet.writeBlank(row, 0);}).toThrow();
        expect(sheet.writeBlank(row, 0, format)).toBe(sheet);

        row++;
    });

    it('sheet.readFormula reads a formula', function() {
        sheet.writeFormula(row, 0,'=SUM(A1:A10)');
        sheet.writeNum(row, 1, 10);

        expect(function() {sheet.readFormula();}).toThrow();
        expect(function() {sheet.readFormula.call({}, row, 0);}).toThrow();

        var formatRef = {};
        expect(function() {sheet.readFormula(row, 1);}).toThrow();
        expect(sheet.readFormula(row, 0)).toBe('SUM(A1:A10)');
        expect(sheet.readFormula(row, 0, formatRef)).toBe('SUM(A1:A10)');
        expect(formatRef.format instanceof format.constructor).toBe(true);

        row++;
    });

    it('sheet.writeComment writes a comment', function() {
        expect(function() {sheet.writeComment();}).toThrow();
        expect(function() {sheet.writeComment.call({}, row, 0, 'comment');}).toThrow();

        expect(sheet.writeComment(row, 0, 'comment')).toBe(sheet);

        row++;
    });

    it('sheet.readComment reads a comment', function() {
        sheet.writeString(row, 0, 'foo');
        sheet.writeComment(row, 0,'comment');

        expect(function() {sheet.readComment();}).toThrow();
        expect(function() {sheet.readComment.call({}, row, 0);}).toThrow();

        var formatRef = {};
        expect(function() {sheet.readComment(row, 1);}).toThrow();
        expect(sheet.readComment(row, 0)).toBe('comment');

        row++;
    });

    it('sheet.readError reads an error', function() {
        sheet.writeStr(row, 0, '');

        expect(function() {sheet.readError();}).toThrow();
        expect(function() {sheet.readError.call({}, row, 0);}).toThrow();

        expect(sheet.readError(row, 0)).toBe(xl.ERRORTYPE_NOERROR);

        row++;
    });


    it('sheet.colWidth reads colum width', function() {
        sheet.setCol(0, 0, 42);

        expect(function() {sheet.colWidth();}).toThrow();
        expect(function() {sheet.colWidth.call({}, 0);}).toThrow();

        expect(sheet.colWidth(0)).toEqual(42);
    });

    it('sheet.rowHeight reads row Height', function() {
        sheet.setRow(1, 42);

        expect(function() {sheet.rowHeight();}).toThrow();
        expect(function() {sheet.rowHeight.call({}, 1);}).toThrow();

        expect(sheet.rowHeight(1)).toEqual(42);
    });

    it('sheet.setCol configures columns', function() {
        expect(function() {sheet.setCol();}).toThrow();
        expect(function() {sheet.setCol.call({}, 0, 0, 42);}).toThrow();
      
        expect(function() {sheet.setCol(0, 0, 42, wrongFormat);}).toThrow();

        expect(sheet.setCol(0, 0, 42)).toBe(sheet);
        expect(sheet.setCol(0, 0, 42, format, false)).toBe(sheet);
    });

    it('sheet.setRow configures rows', function() {
        expect(function() {sheet.setRow();}).toThrow();
        expect(function() {sheet.setRow.call({}, 1, 42);}).toThrow();
      
        expect(function() {sheet.setRow(1, 42, wrongFormat);}).toThrow();

        expect(sheet.setRow(1, 42)).toBe(sheet);
        expect(sheet.setRow(1, 42, format, false)).toBe(sheet);
    });

    it('sheet.rowHidden checks whether a row is hidden', function() {
        expect(function() {sheet.rowHidden();}).toThrow();
        expect(function() {sheet.rowHidden.call({}, row);}).toThrow();

        expect(sheet.rowHidden(row)).toBe(false);
    });

    it('sheet.setRowHidden hides or shows a row', function() {
        expect(function() {sheet.setRowHidden();}).toThrow();
        expect(function() {sheet.setRowHidden.call({}, row, true);}).toThrow();

        expect(sheet.setRowHidden(row, true)).toBe(sheet);
        expect(sheet.rowHidden(row)).toBe(true);
        expect(sheet.setRowHidden(row, false)).toBe(sheet);
        expect(sheet.rowHidden(row)).toBe(false);
    });

    it('sheet.colHidden checks whether a column is hidden', function() {
        expect(function() {sheet.colHidden();}).toThrow();
        expect(function() {sheet.colHidden.call({}, 0);}).toThrow();

        expect(sheet.colHidden(0)).toBe(false);
    });

    it('sheet.setColHidden hides or shows a column', function() {
        expect(function() {sheet.setColHidden();}).toThrow();
        expect(function() {sheet.setColHidden.call({}, 0, true);}).toThrow();

        expect(sheet.setColHidden(0, true)).toBe(sheet);
        expect(sheet.colHidden(0)).toBe(true);
        expect(sheet.setColHidden(0, false)).toBe(sheet);
        expect(sheet.colHidden(0)).toBe(false);
    });
});

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
        var sheet = book.addSheet('foo');
        sheet.writeStr(1, 0, 'bar');

        var file = testUtils.getWriteTestFile();
        shouldThrow(book.writeSync, book, 10);
        shouldThrow(book.writeSync, {}, file);
        expect(book.writeSync(file)).toBe(book);
    });

    it('book.loadSync loads a book in sync mode', function() {
        var file = testUtils.getWriteTestFile();
        shouldThrow(book.loadSync, book, 10);
        shouldThrow(book.loadSync, {}, file);

        expect(book.loadSync(file)).toBe(book);
        var sheet = book.getSheet(0);
        expect(sheet.readStr(1, 0)).toBe('bar');
    });

    it('book.addSheet adds a sheet to a book', function() {
        shouldThrow(book.addSheet, book, 10);
        shouldThrow(book.addSheet, {}, 'foo');
        book.addSheet('baz', 10);

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
        shouldThrow(book.addFormat, {});
        book.addFormat();
        // TODO add a check for format inheritance once format has been
        // completed
    });

    it('book.addFont adds a font', function() {
        shouldThrow(book.addFont, {});
        book.addFont(10);

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
        shouldThrow(book.setTemplate, {});
        expect(book.setTemplate(true)).toBe(book);
    });

    it('book.setKey sets the API key', function() {
        shouldThrow(book.setKey, book, 1, 'a');
        shouldThrow(book.setKey, {}, 'a', 'b');
        expect(book.setKey('a', 'b')).toBe(book);
    });
});
