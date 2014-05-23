var xl = require('../lib/libxl'),
    util = require('util');

function shouldThrow(fun, scope) {
    var args = Array.prototype.slice.call(arguments, 2);

    expect(function() {fun.apply(scope, args);}).toThrow();
}

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
