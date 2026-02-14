const { describe, it, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
const exp = require('constants');

var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The sheet class', function () {
    var book = new xl.Book(xl.BOOK_TYPE_XLS),
        sheet = book.addSheet('foo'),
        format = book.addFormat(),
        wrongBook = new xl.Book(xl.BOOK_TYPE_XLS),
        wrongSheet = wrongBook.addSheet('bar'),
        wrongFormat = wrongBook.addFormat(),
        sheetNameIdx = 0,
        row = 1;

    book.addPicture(testUtils.getTestPicturePath());

    function newSheet() {
        return book.addSheet('foo' + sheetNameIdx++);
    }

    it('sheet.cellType determines cell type', function () {
        sheet.writeStr(row, 0, 'foo').writeNum(row, 1, 10);

        assert.throws(function () {
            sheet.cellType();
        });
        assert.throws(function () {
            sheet.cellType.call({}, row, 0);
        });

        assert.strictEqual(sheet.cellType(row, 0), xl.CELLTYPE_STRING);
        assert.strictEqual(sheet.cellType(row, 1), xl.CELLTYPE_NUMBER);

        row++;
    });

    it('sheet.isFormula checks whether a cell contains a formula', function () {
        sheet.writeStr(row, 0, 'foo').writeFormula(row, 1, '=1');

        assert.throws(function () {
            sheet.isFormula();
        });
        assert.throws(function () {
            sheet.isFormula.call({}, row, 0);
        });

        assert.strictEqual(sheet.isFormula(row, 0), false);
        assert.strictEqual(sheet.isFormula(row, 1), true);

        row++;
    });

    it('sheet.cellFormat retrieves the cell format', function () {
        assert.throws(function () {
            sheet.cellFormat();
        });
        assert.throws(function () {
            sheet.cellFormat.call({}, row, 0);
        });

        var cellFormat = sheet.cellFormat(row, 0);
        assert.strictEqual(format instanceof xl.Format, true);
    });

    it('sheet.setCellFormat sets the cell format', function () {
        assert.throws(function () {
            sheet.setCellFormat();
        });
        assert.throws(function () {
            sheet.setCellFormat.call({}, row, 0, format);
        });

        assert.throws(function () {
            sheet.setCellFormat(row, 0, wrongFormat);
        });
        assert.strictEqual(sheet.setCellFormat(row, 0, format), sheet);
    });

    it('sheet.readStr reads a string', function () {
        sheet.writeStr(row, 0, 'foo');
        sheet.writeNum(row, 1, 10);

        assert.throws(function () {
            sheet.readStr();
        });
        assert.throws(function () {
            sheet.readStr.call({}, row, 0);
        });

        var formatRef = {};
        assert.strictEqual(sheet.readStr(row, 0), 'foo');
        assert.strictEqual(sheet.readStr(row, 0, formatRef), 'foo');
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeStr writes a string', function () {
        assert.throws(function () {
            sheet.writeStr();
        });
        assert.throws(function () {
            sheet.writeStr.call({}, row, 0, 'foo');
        });

        assert.throws(function () {
            sheet.writeStr(row, 0, 'foo', wrongFormat);
        });
        assert.strictEqual(sheet.writeStr(row, 0, 'foo'), sheet);
        assert.strictEqual(sheet.writeStr(row, 0, 'foo', format), sheet);

        row++;
    });

    it('sheet.readRichStr and sheet.writeRichStr read and write a rich string', () => {
        const anotherBook = new xl.Book(xl.BOOK_TYPE_XLSX);

        const format = book.addFormat();
        const anotherFormat = anotherBook.addFormat();

        const richString = book.addRichString();
        const anotherRichString = anotherBook.addRichString();

        shouldThrow(sheet.writeRichStr, sheet, row, 'a', richString);
        shouldThrow(sheet.writeRichStr, {}, row, 0, richString);
        shouldThrow(sheet.writeRichStr, sheet, row, 0, anotherRichString);
        shouldThrow(sheet.writeRichStr, sheet, row, 0, richString, anotherFormat);

        assert.strictEqual(sheet.writeRichStr(row, 0, richString), sheet);
        assert.strictEqual(sheet.writeRichStr(row, 1, richString, format), sheet);

        shouldThrow(sheet.readRichStr, sheet, row, 'a');
        shouldThrow(sheet.readRichStr, {}, row, 0);

        let formatRef = {};
        assert.ok(sheet.readRichStr(row, 0) instanceof xl.RichString);
        assert.ok(sheet.readRichStr(row, 0, formatRef) instanceof xl.RichString);
        assert.ok(formatRef.format instanceof xl.Format);

        formatRef = {};
        assert.ok(sheet.readRichStr(row, 1, formatRef) instanceof xl.RichString);
        assert.ok(formatRef.format instanceof xl.Format);
    });

    it('sheet.readNum reads a number', function () {
        sheet.writeNum(row, 0, 10);

        assert.throws(function () {
            sheet.readNum();
        });
        assert.throws(function () {
            sheet.readNum.call({}, row, 0);
        });

        var formatRef = {};
        assert.deepStrictEqual(sheet.readNum(row, 0), 10);
        assert.deepStrictEqual(sheet.readNum(row, 0, formatRef), 10);
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeNum writes a Number', function () {
        assert.throws(function () {
            sheet.writeNum();
        });
        assert.throws(function () {
            sheet.writeNum.call({}, row, 0, 10);
        });

        assert.throws(function () {
            sheet.writeNum(row, 0, 10, wrongFormat);
        });
        assert.strictEqual(sheet.writeNum(row, 0, 10), sheet);
        assert.strictEqual(sheet.writeNum(row, 0, 10, format), sheet);

        row++;
    });

    it('sheet.readBool reads a bool', function () {
        sheet.writeBool(row, 0, true);

        assert.throws(function () {
            sheet.readBool();
        });
        assert.throws(function () {
            sheet.readBool.call({}, row, 0);
        });

        var formatRef = {};
        assert.deepStrictEqual(sheet.readBool(row, 0), true);
        assert.deepStrictEqual(sheet.readBool(row, 0, formatRef), true);
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeBool writes a bool', function () {
        assert.throws(function () {
            sheet.writeBool();
        });
        assert.throws(function () {
            sheet.writeBool.call({}, row, 0, true);
        });

        assert.throws(function () {
            sheet.writeBool(row, 0, true, wrongFormat);
        });
        assert.strictEqual(sheet.writeBool(row, 0, true), sheet);
        assert.strictEqual(sheet.writeBool(row, 0, true, format), sheet);

        row++;
    });

    it('sheet.readBlank reads a blank cell', function () {
        sheet.writeBlank(row, 0, format);
        sheet.writeNum(row, 1, 10);

        assert.throws(function () {
            sheet.readBlank();
        });
        assert.throws(function () {
            sheet.readBlank.call({}, row, 0);
        });

        var formatRef = {};
        assert.strictEqual(sheet.readBlank(row, 0) instanceof xl.Format, true);
        assert.strictEqual(sheet.readBlank(row, 0, formatRef) instanceof xl.Format, true);
        assert.strictEqual(formatRef.format instanceof xl.Format, true);
        assert.throws(function () {
            sheet.readBlank(row, 1);
        });

        row++;
    });

    it('sheet.writeBlank writes a blank cell', function () {
        assert.throws(function () {
            sheet.writeBlank();
        });
        assert.throws(function () {
            sheet.writeBlank.call({}, row, 0, format);
        });

        assert.throws(function () {
            sheet.writeBlank(row, 0, wrongFormat);
        });
        assert.throws(function () {
            sheet.writeBlank(row, 0);
        });
        assert.strictEqual(sheet.writeBlank(row, 0, format), sheet);

        row++;
    });

    it('sheet.readFormula and sheet.writeFormula read and write a formula', function () {
        shouldThrow(sheet.writeFormula, sheet, row, 'a', '=SUM(A1:A10)');
        shouldThrow(sheet.writeFormula, {}, row, 0, '=SUM(A1:A10)');

        sheet.writeFormula(row, 0, '=SUM(A1:A10)');
        sheet.writeNum(row, 1, 10);

        assert.throws(function () {
            sheet.readFormula();
        });
        assert.throws(function () {
            sheet.readFormula.call({}, row, 0);
        });

        var formatRef = {};
        assert.throws(function () {
            sheet.readFormula(row, 1);
        });
        assert.strictEqual(sheet.readFormula(row, 0), 'SUM(A1:A10)');
        assert.strictEqual(sheet.readFormula(row, 0, formatRef), 'SUM(A1:A10)');
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeFormulaNum write a formula and a numeric value', function () {
        shouldThrow(sheet.writeFormulaNum, sheet, row, 'a', '=SUM(A1:A10)', 1);
        shouldThrow(sheet.writeFormulaNum, {}, row, 0, '=SUM(A1:A10)', 1);

        sheet.writeFormulaNum(row, 0, '=SUM(A1:A10)', 1);

        assert.strictEqual(sheet.readFormula(row, 0), 'SUM(A1:A10)');

        row++;
    });

    it('sheet.writeFormulaStr write a formula and a numeric value', function () {
        shouldThrow(sheet.writeFormulaStr, sheet, row, 'a', '=SUM(A1:A10)', 'a');
        shouldThrow(sheet.writeFormulaStr, {}, row, 0, '=SUM(A1:A10)', 'a');

        sheet.writeFormulaStr(row, 0, '=SUM(A1:A10)', 'a');

        assert.strictEqual(sheet.readFormula(row, 0), 'SUM(A1:A10)');

        row++;
    });

    it('sheet.writeFormulaBool write a formula and a numeric value', function () {
        shouldThrow(sheet.writeFormulaBool, sheet, row, 'a', '=A1>A2', false);
        shouldThrow(sheet.writeFormulaBool, {}, row, 0, '=A1>A2', false);

        sheet.writeFormulaBool(row, 0, '=A1>A2', false);

        assert.strictEqual(sheet.readFormula(row, 0), 'A1>A2');

        row++;
    });

    it('sheet.writeComment writes a comment', function () {
        assert.throws(function () {
            sheet.writeComment();
        });
        assert.throws(function () {
            sheet.writeComment.call({}, row, 0, 'comment');
        });

        assert.strictEqual(sheet.writeComment(row, 0, 'comment'), sheet);

        row++;
    });

    it('sheet.removeComment removes a comment', function () {
        sheet.writeComment(row, 0, 'comment');

        shouldThrow(sheet.removeComment, sheet, row, '1');
        shouldThrow(sheet.removeComment, {}, row, 0);

        sheet.removeComment(row, 0);

        shouldThrow(sheet.readComment, sheet, row, 0);

        row++;
    });

    it('sheet.readComment reads a comment', function () {
        sheet.writeString(row, 0, 'foo');
        sheet.writeComment(row, 0, 'comment');

        assert.throws(function () {
            sheet.readComment();
        });
        assert.throws(function () {
            sheet.readComment.call({}, row, 0);
        });

        var formatRef = {};
        assert.throws(function () {
            sheet.readComment(row, 1);
        });
        assert.strictEqual(sheet.readComment(row, 0), 'comment');

        row++;
    });

    it('sheet.isDate checks whether a cell contains a date', function () {
        sheet.writeNum(row, 0, book.datePack(1980, 8, 19), book.addFormat().setNumFormat(xl.NUMFORMAT_DATE));

        shouldThrow(sheet.isDate, sheet, row, 'a');
        shouldThrow(sheet.isDate, {}, row, 0);

        assert.strictEqual(sheet.isDate(row - 1, 0), false);
        assert.strictEqual(sheet.isDate(row, 0), true);

        row++;
    });

    it('sheet.isRichStr checks whether a cell contains a richt string', () => {
        shouldThrow(sheet.isRichStr, sheet, row, 'a');
        shouldThrow(sheet.isRichStr, {}, row, 0);

        // TODO: add a positive test once rich strings are supported
        assert.strictEqual(sheet.isRichStr(row, 0), false);
    });

    it('sheet.readError reads an error', function () {
        sheet.writeStr(row, 0, '');

        assert.throws(function () {
            sheet.readError();
        });
        assert.throws(function () {
            sheet.readError.call({}, row, 0);
        });

        assert.strictEqual(sheet.readError(row, 0), xl.ERRORTYPE_NOERROR);

        row++;
    });

    it('sheet.writeError writes an error', () => {
        shouldThrow(sheet.writeError, sheet, row, '0', xl.ERRORTYPE_DIV_0);
        shouldThrow(sheet.writeError, {}, row, 0, xl.ERRORTYPE_DIV_0);

        sheet.writeError(row, 0, xl.ERRORTYPE_DIV_0);

        assert.strictEqual(sheet.readError(row, 0), xl.ERRORTYPE_DIV_0);

        row++;
    });

    it('sheet.writeError accepts a format', () => {
        const format = book.addFormat();

        shouldThrow(sheet.writeError, sheet, row, 0, xl.ERRORTYPE_DIV_0, 10);
        shouldThrow(sheet.writeError, {}, row, 0, xl.ERRORTYPE_DIV_0, format);

        sheet.writeError(row, 0, xl.ERRORTYPE_DIV_0, format);

        assert.strictEqual(sheet.readError(row, 0), xl.ERRORTYPE_DIV_0);

        row++;
    });

    it('sheet.colWidth reads colum width', function () {
        sheet.setCol(0, 0, 42);

        assert.throws(function () {
            sheet.colWidth();
        });
        assert.throws(function () {
            sheet.colWidth.call({}, 0);
        });

        assert.deepStrictEqual(sheet.colWidth(0), 42);
    });

    it('sheet.rowHeight reads row Height', function () {
        sheet.setRow(1, 42);

        assert.throws(function () {
            sheet.rowHeight();
        });
        assert.throws(function () {
            sheet.rowHeight.call({}, 1);
        });

        assert.deepStrictEqual(sheet.rowHeight(1), 42);
    });

    it('sheet.colWidthPx reads column width in pixels', () => {
        shouldThrow(sheet.colWidthPx, sheet, '1');
        shouldThrow(sheet.colWidthPx, {}, 1);

        assert.strictEqual(typeof sheet.colWidthPx(1), 'number');
    });

    it('sheet.colWidthPx reads column width in pixels', () => {
        shouldThrow(sheet.rowHeightPx, sheet, '1');
        shouldThrow(sheet.rowHeightPx, {}, 1);

        assert.strictEqual(typeof sheet.rowHeightPx(1), 'number');
    });

    it('sheet.setCol configures columns', function () {
        assert.throws(function () {
            sheet.setCol();
        });
        assert.throws(function () {
            sheet.setCol.call({}, 0, 0, 42);
        });

        assert.throws(function () {
            sheet.setCol(0, 0, 42, wrongFormat);
        });

        assert.strictEqual(sheet.setCol(0, 0, 42), sheet);
        assert.strictEqual(sheet.setCol(0, 0, 42, format, false), sheet);
    });

    it('sheet.setColPx configures columms in pixels', () => {
        shouldThrow(sheet.setColPx, sheet, 0, 0, '10');
        shouldThrow(sheet.setColPx, {}, 0, 0, 10);

        sheet.setColPx(0, 0, 10);

        assert.strictEqual(sheet.colWidthPx(0), 10);
    });

    it('sheet.setColPx accepts format and hidden flag', () => {
        const format = book.addFormat();

        shouldThrow(sheet.setColPx, sheet, 0, 0, 10, format, '1');

        sheet.setColPx(0, 0, 12, format, false);

        assert.strictEqual(sheet.colWidthPx(0), 12);
    });

    it('sheet.setRowPx configures rows in pixels', () => {
        shouldThrow(sheet.setRowPx, sheet, 1, '10');
        shouldThrow(sheet.setRowPx, {}, 1, 10);

        sheet.setRowPx(1, 20);

        assert.strictEqual(sheet.rowHeightPx(1), 20);
    });

    it('sheet.setRowPx accepts format and hidden flag', () => {
        const format = book.addFormat();

        shouldThrow(sheet.setRowPx, sheet, 1, 10, format, '1');

        sheet.setRowPx(1, 12, format, false);

        assert.strictEqual(sheet.rowHeightPx(1), 12);
    });

    it('sheet.setColPx accepts format and hidden flag', () => {
        const format = book.addFormat();

        shouldThrow(sheet.setColPx, {}, 0, 0, 10, format, '1');

        sheet.setColPx(0, 0, 12, format, false);

        assert.strictEqual(sheet.colWidthPx(0), 12);
    });

    it('sheet.setRow configures rows', function () {
        assert.throws(function () {
            sheet.setRow();
        });
        assert.throws(function () {
            sheet.setRow.call({}, 1, 42);
        });

        assert.throws(function () {
            sheet.setRow(1, 42, wrongFormat);
        });

        assert.strictEqual(sheet.setRow(1, 42), sheet);
        assert.strictEqual(sheet.setRow(1, 42, format, false), sheet);
    });

    it('sheet.colFormat retrieves a column format', () => {
        const format = book.addFormat();

        shouldThrow(sheet.colFormat, sheet, '0');
        shouldThrow(sheet.colFormat, {}, 0);

        assert.strictEqual(sheet.colFormat(0) instanceof xl.Format, true);
    });

    it('sheet.rowFormat retrieves a row format', () => {
        const format = book.addFormat();

        shouldThrow(sheet.rowFormat, sheet, '1');
        shouldThrow(sheet.rowFormat, {}, 1);

        assert.strictEqual(sheet.rowFormat(1) instanceof xl.Format, true);
    });

    it('sheet.rowHidden checks whether a row is hidden', function () {
        assert.throws(function () {
            sheet.rowHidden();
        });
        assert.throws(function () {
            sheet.rowHidden.call({}, row);
        });

        assert.strictEqual(sheet.rowHidden(row), false);
    });

    it('sheet.setRowHidden hides or shows a row', function () {
        assert.throws(function () {
            sheet.setRowHidden();
        });
        assert.throws(function () {
            sheet.setRowHidden.call({}, row, true);
        });

        assert.strictEqual(sheet.setRowHidden(row, true), sheet);
        assert.strictEqual(sheet.rowHidden(row), true);
        assert.strictEqual(sheet.setRowHidden(row, false), sheet);
        assert.strictEqual(sheet.rowHidden(row), false);
    });

    it('sheet.colHidden checks whether a column is hidden', function () {
        assert.throws(function () {
            sheet.colHidden();
        });
        assert.throws(function () {
            sheet.colHidden.call({}, 0);
        });

        assert.strictEqual(sheet.colHidden(0), false);
    });

    it('sheet.setColHidden hides or shows a column', function () {
        assert.throws(function () {
            sheet.setColHidden();
        });
        assert.throws(function () {
            sheet.setColHidden.call({}, 0, true);
        });

        assert.strictEqual(sheet.setColHidden(0, true), sheet);
        assert.strictEqual(sheet.colHidden(0), true);
        assert.strictEqual(sheet.setColHidden(0, false), sheet);
        assert.strictEqual(sheet.colHidden(0), false);
    });

    it('sheet.setDefaultRowHeight and sheet.defaultRowHeight control default row height', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.setDefaultRowHeight, sheet, 'a');
        shouldThrow(sheet.setDefaultRowHeight, {}, 66.6);

        assert.strictEqual(sheet.setDefaultRowHeight(66.6), sheet);

        shouldThrow(sheet.defaultRowHeight, {});

        assert.ok(Math.abs(sheet.defaultRowHeight() - 66.6) < testUtils.epsilon);
    });

    it('sheet.setMerge, sheet.getMerge and sheet.delMerge manage merged cells', function () {
        shouldThrow(sheet.getMerge, sheet, row, 0);
        shouldThrow(sheet.delMerge, sheet, row, 0);

        shouldThrow(sheet.setMerge, sheet, 'a', row + 2, 0, 3);
        shouldThrow(sheet.setMerge, {}, row, row + 2, 0, 3);
        assert.strictEqual(sheet.setMerge(row, row + 2, 0, 3), sheet);

        shouldThrow(sheet.getMerge, sheet, 'a', 0);
        shouldThrow(sheet.getMerge, {}, row, 0);
        merge = sheet.getMerge(row, 0);
        assert.strictEqual(merge.rowFirst, row);
        assert.strictEqual(merge.rowLast, row + 2);
        assert.strictEqual(merge.colFirst, 0);
        assert.strictEqual(merge.colLast, 3);

        shouldThrow(sheet.delMerge, sheet, 'a', 0);
        shouldThrow(sheet.delMerge, {}, row, 0);
        sheet.delMerge(row, 0);

        shouldThrow(sheet.delMerge, sheet, row, 0);
    });

    it('sheet.mergeSize, sheet.merge and sheet.delMergeByIndex manage merged cells by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');
        const mergeSizeStart = sheet.mergeSize();

        shouldThrow(sheet.mergeSize, {});
        shouldThrow(sheet.merge, sheet, 'a');
        shouldThrow(sheet.merge, {}, 1);
        shouldThrow(sheet.delMergeByIndex, sheet, 'a');
        shouldThrow(sheet.delMergeByIndex, {}, 1);

        sheet.setMerge(row, row + 2, 0, 3);

        assert.strictEqual(sheet.mergeSize(), mergeSizeStart + 1);
        const mergeIdx = sheet.mergeSize() - 1;

        merge = sheet.merge(mergeIdx);
        assert.strictEqual(merge.rowFirst, row);
        assert.strictEqual(merge.rowLast, row + 2);
        assert.strictEqual(merge.colFirst, 0);
        assert.strictEqual(merge.colLast, 3);

        assert.strictEqual(sheet.delMergeByIndex(mergeIdx), sheet);
        assert.strictEqual(sheet.mergeSize(), mergeSizeStart);
    });

    it('sheet.pictureSize returns the number of pictures on a sheet', function () {
        shouldThrow(sheet.pictureSize, {});
        assert.strictEqual(sheet.pictureSize(), 0);
    });

    it('sheet.setPicture and sheet.getPicture add and inspect pictures', function () {
        shouldThrow(sheet.setPicture, sheet, row, true);
        shouldThrow(sheet.setPicture, {}, row, 0, 0, 1, 0, 0);
        shouldThrow(sheet.setPicture, sheet, row, 0, 0, 1, 0, 0, 'a');

        assert.strictEqual(sheet.setPicture(row, 0, 0, 1, 0, 0), sheet);
        assert.strictEqual(sheet.setPicture(row, 0, 0, 1, 0, 0, xl.POSITION_ONLY_MOVE), sheet);

        var idx = sheet.pictureSize() - 1;
        shouldThrow(sheet.getPicture, sheet, true);
        shouldThrow(sheet.getPicture, {}, idx);

        var pic = sheet.getPicture(idx);
        assert.strictEqual(pic.rowTop, row);
        assert.strictEqual(pic.colLeft, 0);
        assert.strictEqual(pic.width, testUtils.testPictureWidth);
        assert.strictEqual(pic.height, testUtils.testPictureHeight);
        assert.strictEqual(pic.hasOwnProperty('rowBottom'), true);
        assert.strictEqual(pic.hasOwnProperty('colRight'), true);

        row = pic.rowBottom + 1;
    });

    it('ssheet.getPicture can read links', function () {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        book.addPictureAsLink(testUtils.getTestPicturePath());
        sheet.setPicture(row, 0, 0, 1, 0, 0);

        var idx = sheet.pictureSize() - 1;

        var pic = sheet.getPicture(idx);
        assert.strictEqual(typeof pic.linkPath, 'string');

        row = pic.rowBottom + 1;
    });

    it('sheet.setPicture2 adds pictures by width and height instead of scale', function () {
        shouldThrow(sheet.setPicture2, sheet, row, true);
        shouldThrow(sheet.setPicture2, {}, row, 0, 0, 100, 100, 0, 0);
        shouldThrow(sheet.setPicture2, sheet, row, 0, 0, 100, 100, 0, 0, 'a');

        assert.strictEqual(sheet.setPicture2(row, 0, 0, 100, 200, 0, 0), sheet);
        assert.strictEqual(sheet.setPicture2(row, 0, 0, 100, 200, 0, 0, xl.POSITION_ONLY_MOVE), sheet);

        var idx = sheet.pictureSize() - 1;

        var pic = sheet.getPicture(idx);
        assert.strictEqual(pic.rowTop, row);
        assert.strictEqual(pic.colLeft, 0);
        assert.strictEqual(pic.width, 100);
        assert.strictEqual(pic.height, 200);
        assert.strictEqual(pic.hasOwnProperty('rowBottom'), true);
        assert.strictEqual(pic.hasOwnProperty('colRight'), true);

        row = pic.rowBottom + 1;
    });

    it('sheet.removePicture removes a picture from the sheet', () => {
        const pictureSizeOld = sheet.pictureSize();
        sheet.setPicture(row, 0, 0);

        shouldThrow(sheet.removePicture, sheet, row, 'a');
        shouldThrow(sheet.removePicture, {}, row, 0);

        assert.strictEqual(sheet.removePicture(row, 0), sheet);

        assert.strictEqual(sheet.pictureSize(), pictureSizeOld);
    });

    it('sheet.removePictureByIndex removes a picture from the sheet by index', () => {
        const pictureSizeOld = sheet.pictureSize();
        sheet.setPicture(row, 0, 0);

        shouldThrow(sheet.removePictureByIndex, sheet, 'a');
        shouldThrow(sheet.removePictureByIndex, {}, pictureSizeOld - 1);

        assert.strictEqual(sheet.removePictureByIndex(pictureSizeOld - 1), sheet);

        assert.strictEqual(sheet.pictureSize(), pictureSizeOld);
    });

    it(
        'sheet.getHorPageBreak, sheet.setHorPageBreak and sheet.horPageBreakSize' + 'manage horizontal page breaks',
        function () {
            shouldThrow(sheet.getHorPageBreakSize, {});
            var n = sheet.getHorPageBreakSize();

            shouldThrow(sheet.setHorPageBreak, sheet, 'a');
            shouldThrow(sheet.setHorPageBreak, {}, row);
            assert.strictEqual(sheet.setHorPageBreak(row), sheet);
            assert.strictEqual(sheet.getHorPageBreakSize(), n + 1);

            shouldThrow(sheet.getHorPageBreak, sheet, 'a');
            shouldThrow(sheet.getHorPageBreak, {}, n + 1);
            assert.strictEqual(sheet.getHorPageBreak(n), row);

            sheet.setHorPageBreak(row, false);
            assert.strictEqual(sheet.getHorPageBreakSize(), n);
        },
    );

    it(
        'sheet.getVerPageBreak, sheet.setVerPageBreak and sheet.horPageBreakSize' + 'manage vertical page breaks',
        function () {
            shouldThrow(sheet.getVerPageBreakSize, {});
            var n = sheet.getVerPageBreakSize();

            shouldThrow(sheet.setVerPageBreak, sheet, 'a');
            shouldThrow(sheet.setVerPageBreak, {}, row);
            assert.strictEqual(sheet.setVerPageBreak(10), sheet);
            assert.strictEqual(sheet.getVerPageBreakSize(), n + 1);

            shouldThrow(sheet.getVerPageBreak, sheet, 'a');
            shouldThrow(sheet.getVerPageBreak, {}, n + 1);
            assert.strictEqual(sheet.getVerPageBreak(n), 10);

            sheet.setVerPageBreak(10, false);
            assert.strictEqual(sheet.getVerPageBreakSize(), n);
        },
    );

    it('sheet.split splits a sheet', function () {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo');

        shouldThrow(sheet.split, sheet, 'a', 1);
        shouldThrow(sheet.split, {}, 2, 2);
        assert.strictEqual(sheet.split(2, 2), sheet);
    });

    it('sheet.splitInfo obtains info about a sheet split', () => {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo');

        sheet.split(2, 3);

        shouldThrow(sheet.splitInfo, {});

        const result = sheet.splitInfo();

        assert.strictEqual(result.row, 2);
        assert.strictEqual(result.col, 3);
    });

    it('sheet.groupRows groups rows', function () {
        var sheet = newSheet();

        shouldThrow(sheet.groupRows, sheet, 1, 'a');
        shouldThrow(sheet.groupRows, {}, 1, 2);
        assert.strictEqual(sheet.groupRows(1, 2), sheet);
    });

    it('sheet.groupCols groups columns', function () {
        var sheet = newSheet();

        shouldThrow(sheet.groupCols, sheet, 1, 'a');
        shouldThrow(sheet.groupCols, {}, 1, 2);
        assert.strictEqual(sheet.groupCols(1, 2), sheet);
    });

    it('sheet.setGroupSummaryBelow / sheet.groupSummaryBelow manage vert summary position', function () {
        shouldThrow(sheet.setGroupSummaryBelow, sheet, 1);
        shouldThrow(sheet.setGroupSummaryBelow, {}, true);
        assert.strictEqual(sheet.setGroupSummaryBelow(true), sheet);

        shouldThrow(sheet.groupSummaryBelow, {});
        assert.strictEqual(sheet.groupSummaryBelow(), true);

        assert.strictEqual(sheet.setGroupSummaryBelow(false).groupSummaryBelow(), false);
    });

    it('sheet.setGroupSummaryRight / sheet.groupSummaryRight manage hor summary position', function () {
        shouldThrow(sheet.setGroupSummaryRight, sheet, 1);
        shouldThrow(sheet.setGroupSummaryRight, {}, true);
        assert.strictEqual(sheet.setGroupSummaryRight(true), sheet);

        shouldThrow(sheet.groupSummaryRight, {});
        assert.strictEqual(sheet.groupSummaryRight(), true);

        assert.strictEqual(sheet.setGroupSummaryRight(false).groupSummaryRight(), false);
    });

    it('sheet.clear clears the sheet', function () {
        var sheet = newSheet();

        sheet.writeStr(1, 1, 'foo').writeStr(2, 1, 'bar');

        shouldThrow(sheet.clear, sheet, 'a');
        shouldThrow(sheet.clear, {}, 1, 1, 1, 1);
        assert.strictEqual(sheet.clear(1, 1, 1, 1), sheet);

        assert.throws(function () {
            sheet.readStr(1, 1);
        });
        assert.strictEqual(sheet.readStr(2, 1), 'bar');
    });

    it('sheet.insertRow and sheet.insertCol insert rows and cols', function () {
        var sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        shouldThrow(sheet.insertRow, sheet, 'a', 2);
        shouldThrow(sheet.insertRow, {}, 2, 3);
        shouldThrow(sheet.insertCol, sheet, 'a', 2);
        shouldThrow(sheet.insertCol, {}, 2, 3);

        assert.strictEqual(sheet.insertRow(2, 3), sheet);
        assert.strictEqual(sheet.insertCol(2, 3), sheet);

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.insertRow and sheet.insertCol support updateNamedRanges', function () {
        var sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        shouldThrow(sheet.insertRow, sheet, 2, 3, 'a');
        shouldThrow(sheet.insertRow, {}, 2, 3, true);
        shouldThrow(sheet.insertCol, sheet, 2, 3, 'a');
        shouldThrow(sheet.insertCol, {}, 2, 3, true);

        assert.strictEqual(sheet.insertRow(2, 3, true), sheet);
        assert.strictEqual(sheet.insertCol(2, 3, true), sheet);

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.insertRowAsync and sheet.insertColAsync insert rows and cols in async mode', async () => {
        const sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        shouldThrow(sheet.insertRowAsync, sheet, 'a', 2, function () {});
        shouldThrow(sheet.insertRowAsync, {}, 2, 3, function () {});

        const insertRowResult = util.promisify(sheet.insertRowAsync.bind(sheet))(2, 3);
        shouldThrow(book.sheetCount, book);

        await insertRowResult;

        shouldThrow(sheet.insertColAsync, sheet, 'a', 2, function () {});
        shouldThrow(sheet.insertColAsync, {}, 2, 3, function () {});

        const insertColResult = util.promisify(sheet.insertColAsync.bind(sheet))(2, 3);
        shouldThrow(book.sheetCount, book);

        await insertColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.insertRowAsync and sheet.insertColAsync insert rows support updateNamedRanges', async () => {
        const sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        shouldThrow(sheet.insertRowAsync, sheet, 2, 3, 'a', function () {});
        shouldThrow(sheet.insertRowAsync, {}, 2, 3, true, function () {});

        const insertRowResult = util.promisify(sheet.insertRowAsync.bind(sheet))(2, 3, true);
        shouldThrow(book.sheetCount, book);

        await insertRowResult;

        shouldThrow(sheet.insertColAsync, sheet, 2, 3, 'a', function () {});
        shouldThrow(sheet.insertColAsync, {}, 2, 3, true, function () {});

        const insertColResult = util.promisify(sheet.insertColAsync.bind(sheet))(2, 3, true);
        shouldThrow(book.sheetCount, book);

        await insertColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.removeRow and sheet.removeCol remove rows and cols', function () {
        var sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        shouldThrow(sheet.removeRow, sheet, 'a', 2);
        shouldThrow(sheet.removeRow, {}, 2, 3);
        shouldThrow(sheet.removeCol, sheet, 'a', 2);
        shouldThrow(sheet.removeCol, {}, 2, 3);

        assert.strictEqual(sheet.removeRow(2, 3), sheet);
        assert.strictEqual(sheet.removeCol(2, 3), sheet);

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.removeRow and sheet.removeCol support updateNamedRanges', function () {
        var sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        shouldThrow(sheet.removeRow, sheet, 2, 3, 'a');
        shouldThrow(sheet.removeRow, {}, 2, 3, true);
        shouldThrow(sheet.removeCol, sheet, 2, 3, 'a');
        shouldThrow(sheet.removeCol, {}, 2, 3, true);

        assert.strictEqual(sheet.removeRow(2, 3, true), sheet);
        assert.strictEqual(sheet.removeCol(2, 3, true), sheet);

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.removeRowAsync and sheet.removeColAsync remove rows and cols in async mode', async () => {
        const sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        shouldThrow(sheet.removeRowAsync, sheet, 'a', 2, () => undefined);
        shouldThrow(sheet.removeRowAsync, {}, 2, 3, () => undefined);

        const removeRowResult = util.promisify(sheet.removeRowAsync.bind(sheet))(2, 3);
        shouldThrow(book.sheetCount, book);

        await removeRowResult;

        shouldThrow(sheet.removeColAsync, sheet, 'a', 2, () => undefined);
        shouldThrow(sheet.removeColAsync, {}, 2, 3, () => undefined);

        const removeColResult = util.promisify(sheet.removeColAsync.bind(sheet))(2, 3);
        shouldThrow(book.sheetCount, book);

        await removeColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.removeRowAsync and sheet.removeColAsync support updateNamedRanges', async () => {
        const sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        shouldThrow(sheet.removeRowAsync, sheet, 2, 3, 'a', () => undefined);
        shouldThrow(sheet.removeRowAsync, {}, 2, 3, true, () => undefined);

        const removeRowResult = util.promisify(sheet.removeRowAsync.bind(sheet))(2, 3, true);
        shouldThrow(book.sheetCount, book);

        await removeRowResult;

        shouldThrow(sheet.removeColAsync, sheet, 2, 3, 'a', () => undefined);
        shouldThrow(sheet.removeColAsync, {}, 2, 3, true, () => undefined);

        const removeColResult = util.promisify(sheet.removeColAsync.bind(sheet))(2, 3, true);
        shouldThrow(book.sheetCount, book);

        await removeColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.copyCell copies a cell', function () {
        sheet.writeStr(row, 0, 'baz');

        shouldThrow(sheet.copyCell, sheet, row, 0, row, 'a');
        shouldThrow(sheet.copyCell, {}, row, 0, row, 1);

        assert.strictEqual(sheet.copyCell(row, 0, row, 1), sheet);
        assert.strictEqual(sheet.readStr(row, 1), 'baz');

        row++;
    });

    it('sheet.firstRow, sheet.firstCol, sheet.lastRow, sheet.lastCol return ' + 'the spreadsheet limits', function () {
        var sheet = newSheet();

        sheet.writeNum(2, 1, 1).writeNum(5, 5, 1);

        shouldThrow(sheet.firstRow, {});
        shouldThrow(sheet.firstCol, {});
        shouldThrow(sheet.lastRow, {});
        shouldThrow(sheet.lastCol, {});

        assert.strictEqual(sheet.firstRow(), 2);
        assert.strictEqual(sheet.firstCol(), 1);
        assert.strictEqual(sheet.lastRow(), 6);
        assert.strictEqual(sheet.lastCol(), 6);
    });

    it(
        'sheet.firstFilledRow, sheet.firstFilledCol, sheet.lastFilledRow, sheet.lastFilledCol return ' +
            'the spreadsheet limits',
        function () {
            var sheet = newSheet();

            sheet.writeNum(2, 1, 1).writeNum(5, 5, 1);

            shouldThrow(sheet.firstFilledRow, {});
            shouldThrow(sheet.firstFilledCol, {});
            shouldThrow(sheet.lastFilledRow, {});
            shouldThrow(sheet.lastFilledCol, {});

            assert.strictEqual(sheet.firstFilledRow(), 0);
            assert.strictEqual(sheet.firstFilledCol(), 0);
            assert.strictEqual(sheet.lastFilledRow(), 6);
            assert.strictEqual(sheet.lastFilledCol(), 6);
        },
    );

    it('sheet.displayGridlines and sheet.setDisplayGridlines manage display of gridlines', function () {
        shouldThrow(sheet.setDisplayGridlines, sheet, 1);
        shouldThrow(sheet.setDisplayGridlines, {});
        assert.strictEqual(sheet.setDisplayGridlines(), sheet);

        shouldThrow(sheet.displayGridlines, {});
        assert.strictEqual(sheet.displayGridlines(), true);
        assert.strictEqual(sheet.setDisplayGridlines(false).displayGridlines(), false);
    });

    it('sheet.printGridlines and sheet.setPrintGridlines manage print of gridlines', function () {
        shouldThrow(sheet.setPrintGridlines, sheet, 1);
        shouldThrow(sheet.setPrintGridlines, {});
        assert.strictEqual(sheet.setPrintGridlines(), sheet);

        shouldThrow(sheet.printGridlines, {});
        assert.strictEqual(sheet.printGridlines(), true);
        assert.strictEqual(sheet.setPrintGridlines(false).printGridlines(), false);
    });

    it('sheet.zoom and sheet.setZoom control sheet zoom', function () {
        shouldThrow(sheet.setZoom, sheet, true);
        shouldThrow(sheet.setZoom, {}, 10);
        assert.strictEqual(sheet.setZoom(10), sheet);

        shouldThrow(sheet.zoom, {});
        assert.strictEqual(sheet.zoom(), 10);
        assert.strictEqual(sheet.setZoom(1).zoom(), 1);
    });

    it('sheet.setLandscape and sheet.landscape control landscape mode', function () {
        shouldThrow(sheet.setLandscape, sheet, 1);
        shouldThrow(sheet.setLandscape, {});
        assert.strictEqual(sheet.setLandscape(), sheet);

        shouldThrow(sheet.landscape, {});
        assert.strictEqual(sheet.landscape(), true);
        assert.strictEqual(sheet.setLandscape(false).landscape(), false);
    });

    it('sheet.paper and sheet.setPaper control paper format', function () {
        shouldThrow(sheet.setPaper, sheet, true);
        shouldThrow(sheet.setPaper, {}, xl.PAPER_NOTE);
        assert.strictEqual(sheet.setPaper(xl.PAPER_NOTE), sheet);

        shouldThrow(sheet.paper, {});
        assert.strictEqual(sheet.paper(), xl.PAPER_NOTE);
        assert.strictEqual(sheet.setPaper().paper(), xl.PAPER_DEFAULT);
    });

    it('sheet.header, sheet.setHeader and sheet.headerMargin control the header', function () {
        shouldThrow(sheet.setHeader, sheet, 1);
        shouldThrow(sheet.setHeader, {}, 'fuppe');
        assert.strictEqual(sheet.setHeader('fuppe'), sheet);

        shouldThrow(sheet.header, {});
        assert.strictEqual(sheet.header(), 'fuppe');

        shouldThrow(sheet.headerMargin, {});
        assert.strictEqual(sheet.headerMargin(), 0.5);
        assert.strictEqual(sheet.setHeader('fuppe', 11.11).headerMargin(), 11.11);
    });

    it('sheet.footer, sheet.setFooter and sheet.footerMargin control the footer', function () {
        shouldThrow(sheet.setFooter, sheet, 1);
        shouldThrow(sheet.setFooter, {}, 'foppe');
        assert.strictEqual(sheet.setFooter('foppe'), sheet);

        shouldThrow(sheet.footer, {});
        assert.strictEqual(sheet.footer(), 'foppe');

        shouldThrow(sheet.footerMargin, {});
        assert.strictEqual(sheet.footerMargin(), 0.5);
        assert.strictEqual(sheet.setFooter('foppe', 11.11).footerMargin(), 11.11);
    });

    ['Left', 'Right', 'Bottom', 'Top'].forEach(function (side) {
        var getter = 'margin' + side,
            setter = 'setMargin' + side;

        it(util.format('sheet.%s and sheet.%s control %s margin', getter, setter, side.toLowerCase()), function () {
            shouldThrow(sheet[setter], sheet, true);
            shouldThrow(sheet[setter], {}, 11.11);
            assert.strictEqual(sheet[setter](11.11), sheet);

            shouldThrow(sheet[getter], {});
            assert.strictEqual(sheet[getter](), 11.11);
            assert.strictEqual(sheet[setter](12.12)[getter](), 12.12);
        });
    });

    it('sheet.printRowCol and sheet.setPrintRowCol control printing of row / col headers', function () {
        shouldThrow(sheet.setPrintRowCol, sheet, 1);
        shouldThrow(sheet.setPrintRowCol, {});
        assert.strictEqual(sheet.setPrintRowCol(), sheet);

        shouldThrow(sheet.printRowCol, {});
        assert.strictEqual(sheet.printRowCol(), true);
        assert.strictEqual(sheet.setPrintRowCol(false).printRowCol(), false);
    });

    it(
        'sheet.setPrintRepeatRows / sheet.printRepeatRows, sheet.setPrintRepeatCols / sheet.printRepeatCols, sheet.clearPrintRepeat ' +
            ' control row / col repeat in print',
        function () {
            const book = new xl.Book(xl.BOOK_TYPE_XLS);
            const sheet = book.addSheet('foo');
            sheet.writeStr(1, 1, 'foo');
            sheet.writeStr(2, 2, 'foo');

            shouldThrow(sheet.setPrintRepeatRows, sheet, true, 1);
            shouldThrow(sheet.setPrintRepeatRows, {}, 1, 2);
            assert.strictEqual(sheet.setPrintRepeatRows(1, 2), sheet);

            shouldThrow(sheet.printRepeatRows, {});
            const printRepeatRows = sheet.printRepeatRows();

            assert.strictEqual(printRepeatRows.rowFirst, 1);
            assert.strictEqual(printRepeatRows.rowLast, 2);

            shouldThrow(sheet.setPrintRepeatCols, sheet, true, 1);
            shouldThrow(sheet.setPrintRepeatCols, {}, 1, 2);
            assert.strictEqual(sheet.setPrintRepeatCols(1, 2), sheet);

            shouldThrow(sheet.printRepeatCols, {});
            const printRepeatCols = sheet.printRepeatCols();

            assert.strictEqual(printRepeatCols.colFirst, 1);
            assert.strictEqual(printRepeatCols.colLast, 2);

            shouldThrow(sheet.clearPrintRepeats, {});
            assert.strictEqual(sheet.clearPrintRepeats(), sheet);

            shouldThrow(sheet.printRepeatCols, sheet);
            shouldThrow(sheet.printRepeatRows, sheet);
        },
    );

    it('sheet.setPrintArea and sheet.clearPrintArea manage the print area', function () {
        shouldThrow(sheet.setPrintArea, sheet, true, 1, 5, 5);
        shouldThrow(sheet.setPrintArea, {}, 1, 1, 5, 5);
        assert.strictEqual(sheet.setPrintArea(1, 1, 5, 5), sheet);

        shouldThrow(sheet.clearPrintArea, {});
        assert.strictEqual(sheet.clearPrintArea(), sheet);
    });

    it(
        'sheet.getNamedRange, sheet.setNamedRange, sheet.delNamedRange, sheet.namedRangeSize and sheet.namedRange ' +
            'manage named ranges',
        function () {
            var sheet2 = newSheet();

            function assertRange(range, rowFirst, rowLast, colFirst, colLast, name, scope) {
                assert.strictEqual(range.rowFirst, rowFirst);
                assert.strictEqual(range.rowLast, rowLast);
                assert.strictEqual(range.colFirst, colFirst);
                assert.strictEqual(range.colLast, colLast);
                if (typeof name !== 'undefined') {
                    assert.strictEqual(range.name, name);
                }
                if (typeof scope !== 'undefined') {
                    assert.strictEqual(range.scopeId, scope);
                }
            }

            shouldThrow(sheet.setNamedRange, sheet, 'range1', 'a', 1, 5, 5);
            shouldThrow(sheet.setNamedRange, {}, 'range1', 1, 1, 5, 5);
            assert.strictEqual(sheet.setNamedRange('range1', 1, 1, 5, 5, xl.SCOPE_WORKBOOK), sheet);
            assert.strictEqual(sheet2.setNamedRange('range2', 2, 2, 6, 6), sheet2);

            shouldThrow(sheet.namedRangeSize, {});
            assert.strictEqual(sheet.namedRangeSize(), 1);
            assert.strictEqual(sheet2.namedRangeSize(), 1);

            var range1, range2;
            shouldThrow(sheet.getNamedRange, sheet, 1);
            shouldThrow(sheet.getNamedRange, {}, 'range1');
            shouldThrow(sheet2.getNamedRange, sheet, 'range2', xl.SCOPE_WORKBOOK);
            assertRange(sheet.getNamedRange('range1'), 1, 1, 5, 5);
            assertRange(sheet2.getNamedRange('range2'), 2, 2, 6, 6);

            shouldThrow(sheet.namedRange, sheet, true);
            shouldThrow(sheet.namedRange, {}, 0);
            assertRange(sheet.namedRange(0), 1, 1, 5, 5, 'range1', xl.SCOPE_WORKBOOK);
            assertRange(sheet2.namedRange(0), 2, 2, 6, 6, 'range2');

            shouldThrow(sheet.delNamedRange, sheet, 1);
            shouldThrow(sheet.delNamedRange, {}, 'range1');
            assert.strictEqual(sheet.delNamedRange('range1'), sheet);
            assert.strictEqual(sheet.namedRangeSize(), 0);
        },
    );

    it('sheet.tableSize gets the number of tables on the sheet', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(testUtils.getXlsxTableFile());

        const sheet = book.getSheet(0);

        shouldThrow(sheet.tableSize, {});

        assert.strictEqual(sheet.tableSize(), 1);
    });

    it('sheet.table retrieves a table by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(testUtils.getXlsxTableFile());

        const sheet = book.getSheet(0);

        shouldThrow(sheet.table, sheet, 'a');
        shouldThrow(sheet.table, {}, 0);

        assert.deepStrictEqual(sheet.table(0), { rowFirst: 1, rowLast: 10, colFirst: 1, colLast: 4, totalRowCount: 1 });
    });

    it('sheet.getTable retrieves a table by nyme', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(testUtils.getXlsxTableFile());

        const sheet = book.getSheet(0);

        shouldThrow(sheet.getTable, sheet, 1);
        shouldThrow(sheet.getTable, {}, 'Table');

        assert.deepStrictEqual(sheet.getTable('Table'), {
            rowFirst: 1,
            rowLast: 10,
            colFirst: 1,
            colLast: 4,
            totalRowCount: 1,
        });
    });

    it('sheet.addTable adds a table and returns a Table instance', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.addTable, {});
        shouldThrow(sheet.addTable, sheet, 'T', 0, 5, 0, 'a');

        assert.ok(sheet.addTable('Table1', 0, 5, 0, 3) instanceof xl.Table);
    });

    it('sheet.getTableByName retrieves a table by name', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');
        sheet.addTable('MyTable', 0, 5, 0, 3);

        shouldThrow(sheet.getTableByName, sheet, 1);
        shouldThrow(sheet.getTableByName, {}, 'MyTable');

        assert.ok(sheet.getTableByName('MyTable') instanceof xl.Table);
    });

    it('sheet.getTableByIndex retrieves a table by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');
        sheet.addTable('MyTable', 0, 5, 0, 3);

        shouldThrow(sheet.getTableByIndex, sheet, 'a');
        shouldThrow(sheet.getTableByIndex, {}, 0);

        assert.ok(sheet.getTableByIndex(0) instanceof xl.Table);
    });

    it('sheet.name and sheet.setName manage the sheet name', function () {
        var sheet = book.addSheet('bazzaraz');

        shouldThrow(sheet.name, {});
        assert.strictEqual(sheet.name(), 'bazzaraz');

        shouldThrow(sheet.setName, sheet, 1);
        shouldThrow(sheet.setName, {}, 'fooinator');
        assert.strictEqual(sheet.setName('fooinator'), sheet);
        assert.strictEqual(sheet.name(), 'fooinator');
    });

    it('sheet.protect and sheet.setProtect manage the protection flag', function () {
        shouldThrow(sheet.setProtect, sheet, 1);
        shouldThrow(sheet.setProtect, {});
        assert.strictEqual(sheet.setProtect(), sheet);

        shouldThrow(sheet.protect, {});
        assert.strictEqual(sheet.protect(), true);
        assert.strictEqual(sheet.setProtect(false).protect(), false);
    });

    it('sheet.hidden and sheet.setHidden control wether a sheet is visible', function () {
        shouldThrow(sheet.setHidden, sheet, true);
        shouldThrow(sheet.setHidden, {});
        assert.strictEqual(sheet.setHidden(), sheet);

        shouldThrow(sheet.hidden, {});
        assert.strictEqual(sheet.hidden(), xl.SHEETSTATE_HIDDEN);
        assert.strictEqual(sheet.setHidden(xl.SHEETSTATE_VISIBLE).hidden(), xl.SHEETSTATE_VISIBLE);
    });

    it('sheet.getTopLeftView and sheet.setTopLeftView manage top left visible edge of the sheet', function () {
        var sheet = newSheet();

        shouldThrow(sheet.setTopLeftView, sheet, 5, true);
        shouldThrow(sheet.setTopLeftView, {}, 5, 6);
        assert.strictEqual(sheet.setTopLeftView(5, 6), sheet);

        shouldThrow(sheet.getTopLeftView, {});
        var result = sheet.getTopLeftView();
        assert.strictEqual(result.row, 5);
        assert.strictEqual(result.col, 6);
    });

    it('sheet.addrToRowCol translates a string address to row / col', function () {
        shouldThrow(sheet.addrToRowCol, sheet, 1);
        shouldThrow(sheet.addrToRowCol, {}, 'A1');

        var result = sheet.addrToRowCol('A1');
        assert.strictEqual(result.row, 0);
        assert.strictEqual(result.col, 0);
        assert.strictEqual(result.rowRelative, true);
        assert.strictEqual(result.colRelative, true);
    });

    it('sheet.rowColToAddr builds an address string', function () {
        shouldThrow(sheet.rowColToAddr, sheet, 0, 'a');
        shouldThrow(sheet.rowColToAddr, {}, 0, 0);
        assert.strictEqual(sheet.rowColToAddr(0, 0), 'A1');
        assert.strictEqual(sheet.rowColToAddr(0, 0, false, false), '$A$1');
    });

    it('the hyperlink family of functions manages hyperlinks', () => {
        shouldThrow(sheet.hyperlinkSize, {});
        assert.strictEqual(sheet.hyperlinkSize(), 0);

        shouldThrow(sheet.addHyperlink, sheet, 1, row, row + 2, 1, 3);
        shouldThrow(sheet.addHyperlink, {}, 'http://foo.bar', row, 1, row + 2, 3);

        assert.strictEqual(sheet.addHyperlink('http://foo.bar', row, row + 2, 1, 3), sheet);
        assert.strictEqual(sheet.hyperlinkSize(), 1);

        shouldThrow(sheet.hyperlinkIndex, 'a', 1);
        shouldThrow(sheet.hyperlinkIndex, {}, row, 4);

        assert.strictEqual(sheet.hyperlinkIndex(row, 4), -1);
        assert.strictEqual(sheet.hyperlinkIndex(row, 1), 1);

        shouldThrow(sheet.hyperlink, sheet, 'a');
        shouldThrow(sheet.hyperlink, {}, 0);

        assert.deepStrictEqual(sheet.hyperlink(0), {
            colFirst: 1,
            colLast: 3,
            rowFirst: row,
            rowLast: row + 2,
            hyperlink: 'http://foo.bar',
        });

        shouldThrow(sheet.delHyperlink, sheet, 'a');
        shouldThrow(sheet.delHyperlink, {}, 0);

        assert.strictEqual(sheet.delHyperlink(0), sheet);
        assert.strictEqual(sheet.hyperlinkSize(), 0);

        row++;
    });

    it('the autofilter family of functions manages the auto filter', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.isAutoFilter, {});

        assert.strictEqual(sheet.isAutoFilter(), false);

        shouldThrow(sheet.autoFilter, {});

        assert.ok(sheet.autoFilter() instanceof xl.AutoFilter);

        shouldThrow(sheet.applyFilter, {});

        assert.strictEqual(sheet.applyFilter(), sheet);
        assert.strictEqual(sheet.isAutoFilter(), true);

        shouldThrow(sheet.removeFilter, {});

        assert.strictEqual(sheet.removeFilter(), sheet);
        assert.strictEqual(sheet.isAutoFilter(), false);
    });

    it('sheet.setAutoFitArea sets the auto fit area', () => {
        shouldThrow(sheet.setAutoFitArea, sheet, 1, 1, 10, 'a');

        assert.strictEqual(sheet.setAutoFitArea(1, 1, 10, 10), sheet);
    });

    it('sheet.setTabColor, sheet.getTabColor and sheet.tabColor manage the tab color', () => {
        shouldThrow(sheet.setTabColor, sheet, 'a');
        shouldThrow(sheet.setTabColor, {}, xl.COLOR_BLACK);
        shouldThrow(sheet.setTabColor, sheet, 255, 200, 'a');
        shouldThrow(sheet.setTabColor, {}, 255, 200, 180);
        shouldThrow(sheet.tabColor, {});

        assert.strictEqual(sheet.setTabColor(xl.COLOR_BLACK), sheet);
        assert.strictEqual(sheet.tabColor(), xl.COLOR_BLACK);

        assert.strictEqual(sheet.setTabColor(255, 200, 180), sheet);
        assert.deepStrictEqual(sheet.getTabColor(), { red: 255, green: 200, blue: 180 });
    });

    it('sheet.addIgnoredError configures the ignored error class for a range of cells', () => {
        shouldThrow(sheet.addIgnoredError, sheet, row, 1, row + 2, 3, 'a');
        shouldThrow(sheet.addIgnoredError, {}, row, 1, row + 2, 3, xl.IERR_EVAL_ERROR);

        assert.strictEqual(sheet.addIgnoredError(row, 1, row + 2, 3, xl.IERR_TWODIG_TEXTYEAR), sheet);
    });

    it('sheet.addDataValidation adds a string data validation for a range of cell', () => {
        shouldThrow(
            sheet.addDataValidation,
            sheet,
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            'a',
            'b',
            false,
            false,
            false,
            false,
            'hanni',
            'nanni',
            'fanni',
            'wanni',
            'a',
        );

        shouldThrow(
            sheet.addDataValidation,
            {},
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            'a',
            'b',
            false,
            false,
            false,
            false,
            'hanni',
            'nanni',
            'fanni',
            'wanni',
            xl.VALIDATION_ERRSTYLE_STOP,
        );

        shouldThrow(sheet.addDataValidation, sheet, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 1);
        shouldThrow(sheet.addDataValidation, {}, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 'a');
        shouldThrow(sheet.addDataValidation, sheet, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5);

        assert.strictEqual(
            sheet.addDataValidation(
                xl.VALIDATION_TYPE_WHOLE,
                xl.VALIDATION_OP_EQUAL,
                1,
                5,
                1,
                5,
                'a',
                'b',
                false,
                false,
                false,
                false,
                'hanni',
                'nanni',
                'fanni',
                'wanni',
                xl.VALIDATION_ERRSTYLE_STOP,
            ),
            sheet,
        );

        assert.strictEqual(
            sheet.addDataValidation(xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 'a', 'b'),
            sheet,
        );
    });

    it('sheet.addDataValidationDouble adds a double data validation for a range of cell', () => {
        shouldThrow(
            sheet.addDataValidationDouble,
            sheet,
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            1,
            2,
            false,
            false,
            false,
            false,
            'hanni',
            'nanni',
            'fanni',
            'wanni',
            'a',
        );

        shouldThrow(
            sheet.addDataValidationDouble,
            {},
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            1,
            2,
            false,
            false,
            false,
            false,
            'hanni',
            'nanni',
            'fanni',
            'wanni',
            xl.VALIDATION_ERRSTYLE_STOP,
        );

        shouldThrow(
            sheet.addDataValidationDouble,
            sheet,
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            1,
            'a',
        );
        shouldThrow(
            sheet.addDataValidationDouble,
            {},
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            1,
            2,
        );
        shouldThrow(
            sheet.addDataValidationDouble,
            sheet,
            xl.VALIDATION_TYPE_WHOLE,
            xl.VALIDATION_OP_EQUAL,
            1,
            5,
            1,
            5,
            1,
        );

        assert.strictEqual(
            sheet.addDataValidationDouble(
                xl.VALIDATION_TYPE_WHOLE,
                xl.VALIDATION_OP_EQUAL,
                1,
                5,
                1,
                5,
                1,
                2,
                false,
                false,
                false,
                false,
                'hanni',
                'nanni',
                'fanni',
                'wanni',
                xl.VALIDATION_ERRSTYLE_STOP,
            ),
            sheet,
        );

        assert.strictEqual(
            sheet.addDataValidationDouble(xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 1, 2),
            sheet,
        );
    });

    it('sheet.removeDataValidations removes all data validations', () => {
        assert.strictEqual(sheet.removeDataValidations(), sheet);
    });

    it('sheet.formControlSize returns the number of form controls', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(testUtils.getXlsmFormControlFile());

        const sheet = book.getSheet(0);

        shouldThrow(sheet.formControlSize, {});

        assert.strictEqual(sheet.formControlSize(), 1);
    });

    it('sheet.formControl retrieves a form control by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(testUtils.getXlsmFormControlFile());

        const sheet = book.getSheet(0);

        shouldThrow(sheet.formControl, sheet, 'a');
        shouldThrow(sheet.formControl, {}, 0);

        assert.ok(sheet.formControl(0) instanceof xl.FormControl);
    });

    it('sheet.addConditionalFormatting adds a conditional formatting rule', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.addConditionalFormatting, {});
        shouldThrow(sheet.addConditionalFormatting, sheet, 0, 0, 1, 'a');

        assert.ok(sheet.addConditionalFormatting(0, 0, 1, 1) instanceof xl.ConditionalFormatting);
    });

    it('sheet.getActiveSell and sheet.setActiveCell manage the active cell', () => {
        shouldThrow(sheet.setActiveCell, sheet, 1, 'a');
        shouldThrow(sheet.setActiveCell, {}, 1, 1);

        assert.strictEqual(sheet.setActiveCell(1, 1), sheet);

        shouldThrow(sheet.getActiveCell, {});

        assert.deepStrictEqual(sheet.getActiveCell(), { row: 1, col: 1 });
    });

    it('sheet.addSectionRange, sheet.removeSelection and sheet.selectionRange manage the selection range', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.removeSelection, {});
        assert.strictEqual(sheet.removeSelection(), sheet);

        shouldThrow(sheet.addSelectionRange, sheet, 1);
        shouldThrow(sheet.addSelectionRange, {}, 'A1:A12');

        assert.strictEqual(sheet.addSelectionRange('A2:A10'), sheet);

        shouldThrow(sheet.selectionRange, {});

        assert.strictEqual(sheet.selectionRange(), 'A2:A10');
    });

    it('sheet.setBorder draws borders in an area of cells', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.setBorder, sheet, 0, 0, 0, 0, 'a', 0);
        shouldThrow(sheet.setBorder, {}, 0, 0, 0, 0, xl.BORDERSTYLE_THIN, xl.COLOR_BLACK);

        assert.strictEqual(sheet.setBorder(0, 5, 0, 5, xl.BORDERSTYLE_THIN, xl.COLOR_BLACK), sheet);
    });

    it('sheet.applyFilter2 applies a filter using an AutoFilter object', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        const af = sheet.autoFilter();

        shouldThrow(sheet.applyFilter2, sheet, 1);
        shouldThrow(sheet.applyFilter2, {}, af);

        assert.strictEqual(sheet.applyFilter2(af), sheet);
    });

    it('sheet.conditionalFormattingSize, sheet.conditionalFormatting and sheet.removeConditionalFormatting manage conditional formatting', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        shouldThrow(sheet.conditionalFormattingSize, {});
        assert.strictEqual(sheet.conditionalFormattingSize(), 0);

        sheet.addConditionalFormatting(0, 0, 1, 1);
        assert.strictEqual(sheet.conditionalFormattingSize(), 1);

        shouldThrow(sheet.conditionalFormatting, sheet, 'a');
        shouldThrow(sheet.conditionalFormatting, {}, 0);

        assert.ok(sheet.conditionalFormatting(0) instanceof xl.ConditionalFormatting);

        shouldThrow(sheet.removeConditionalFormatting, sheet, 'a');
        shouldThrow(sheet.removeConditionalFormatting, {}, 0);

        assert.strictEqual(sheet.removeConditionalFormatting(0), sheet);
        assert.strictEqual(sheet.conditionalFormattingSize(), 0);
    });
});
