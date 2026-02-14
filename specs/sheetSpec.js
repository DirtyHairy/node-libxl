import { describe, it, beforeEach } from 'node:test';
import assert from 'node:assert/strict';
import util from 'util';
import xl from '../lib/libxl.js';
import {
    getTestPicturePath,
    testPictureWidth,
    testPictureHeight,
    getXlsxTableFile,
    getXlsmFormControlFile,
    epsilon,
} from './testUtils.js';

describe('The sheet class', () => {
    let book = new xl.Book(xl.BOOK_TYPE_XLS),
        sheet = book.addSheet('foo'),
        format = book.addFormat(),
        wrongBook = new xl.Book(xl.BOOK_TYPE_XLS),
        wrongSheet = wrongBook.addSheet('bar'),
        wrongFormat = wrongBook.addFormat(),
        sheetNameIdx = 0,
        row = 1;

    book.addPicture(getTestPicturePath());

    const newSheet = () => {
        return book.addSheet('foo' + sheetNameIdx++);
    };

    it('sheet.cellType determines cell type', () => {
        sheet.writeStr(row, 0, 'foo').writeNum(row, 1, 10);

        assert.throws(() => {
            sheet.cellType();
        });
        assert.throws(() => {
            sheet.cellType.call({}, row, 0);
        });

        assert.strictEqual(sheet.cellType(row, 0), xl.CELLTYPE_STRING);
        assert.strictEqual(sheet.cellType(row, 1), xl.CELLTYPE_NUMBER);

        row++;
    });

    it('sheet.isFormula checks whether a cell contains a formula', () => {
        sheet.writeStr(row, 0, 'foo').writeFormula(row, 1, '=1');

        assert.throws(() => {
            sheet.isFormula();
        });
        assert.throws(() => {
            sheet.isFormula.call({}, row, 0);
        });

        assert.strictEqual(sheet.isFormula(row, 0), false);
        assert.strictEqual(sheet.isFormula(row, 1), true);

        row++;
    });

    it('sheet.cellFormat retrieves the cell format', () => {
        assert.throws(() => {
            sheet.cellFormat();
        });
        assert.throws(() => {
            sheet.cellFormat.call({}, row, 0);
        });

        const cellFormat = sheet.cellFormat(row, 0);
        assert.strictEqual(format instanceof xl.Format, true);
    });

    it('sheet.setCellFormat sets the cell format', () => {
        assert.throws(() => {
            sheet.setCellFormat();
        });
        assert.throws(() => {
            sheet.setCellFormat.call({}, row, 0, format);
        });

        assert.throws(() => {
            sheet.setCellFormat(row, 0, wrongFormat);
        });
        assert.strictEqual(sheet.setCellFormat(row, 0, format), sheet);
    });

    it('sheet.readStr reads a string', () => {
        sheet.writeStr(row, 0, 'foo');
        sheet.writeNum(row, 1, 10);

        assert.throws(() => {
            sheet.readStr();
        });
        assert.throws(() => {
            sheet.readStr.call({}, row, 0);
        });

        const formatRef = {};
        assert.strictEqual(sheet.readStr(row, 0), 'foo');
        assert.strictEqual(sheet.readStr(row, 0, formatRef), 'foo');
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeStr writes a string', () => {
        assert.throws(() => {
            sheet.writeStr();
        });
        assert.throws(() => {
            sheet.writeStr.call({}, row, 0, 'foo');
        });

        assert.throws(() => {
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

        assert.throws(() => sheet.writeRichStr.call(sheet, row, 'a', richString));
        assert.throws(() => sheet.writeRichStr.call({}, row, 0, richString));
        assert.throws(() => sheet.writeRichStr.call(sheet, row, 0, anotherRichString));
        assert.throws(() => sheet.writeRichStr.call(sheet, row, 0, richString, anotherFormat));

        assert.strictEqual(sheet.writeRichStr(row, 0, richString), sheet);
        assert.strictEqual(sheet.writeRichStr(row, 1, richString, format), sheet);

        assert.throws(() => sheet.readRichStr.call(sheet, row, 'a'));
        assert.throws(() => sheet.readRichStr.call({}, row, 0));

        let formatRef = {};
        assert.ok(sheet.readRichStr(row, 0) instanceof xl.RichString);
        assert.ok(sheet.readRichStr(row, 0, formatRef) instanceof xl.RichString);
        assert.ok(formatRef.format instanceof xl.Format);

        formatRef = {};
        assert.ok(sheet.readRichStr(row, 1, formatRef) instanceof xl.RichString);
        assert.ok(formatRef.format instanceof xl.Format);
    });

    it('sheet.readNum reads a number', () => {
        sheet.writeNum(row, 0, 10);

        assert.throws(() => {
            sheet.readNum();
        });
        assert.throws(() => {
            sheet.readNum.call({}, row, 0);
        });

        const formatRef = {};
        assert.deepStrictEqual(sheet.readNum(row, 0), 10);
        assert.deepStrictEqual(sheet.readNum(row, 0, formatRef), 10);
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeNum writes a Number', () => {
        assert.throws(() => {
            sheet.writeNum();
        });
        assert.throws(() => {
            sheet.writeNum.call({}, row, 0, 10);
        });

        assert.throws(() => {
            sheet.writeNum(row, 0, 10, wrongFormat);
        });
        assert.strictEqual(sheet.writeNum(row, 0, 10), sheet);
        assert.strictEqual(sheet.writeNum(row, 0, 10, format), sheet);

        row++;
    });

    it('sheet.readBool reads a bool', () => {
        sheet.writeBool(row, 0, true);

        assert.throws(() => {
            sheet.readBool();
        });
        assert.throws(() => {
            sheet.readBool.call({}, row, 0);
        });

        const formatRef = {};
        assert.deepStrictEqual(sheet.readBool(row, 0), true);
        assert.deepStrictEqual(sheet.readBool(row, 0, formatRef), true);
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeBool writes a bool', () => {
        assert.throws(() => {
            sheet.writeBool();
        });
        assert.throws(() => {
            sheet.writeBool.call({}, row, 0, true);
        });

        assert.throws(() => {
            sheet.writeBool(row, 0, true, wrongFormat);
        });
        assert.strictEqual(sheet.writeBool(row, 0, true), sheet);
        assert.strictEqual(sheet.writeBool(row, 0, true, format), sheet);

        row++;
    });

    it('sheet.readBlank reads a blank cell', () => {
        sheet.writeBlank(row, 0, format);
        sheet.writeNum(row, 1, 10);

        assert.throws(() => {
            sheet.readBlank();
        });
        assert.throws(() => {
            sheet.readBlank.call({}, row, 0);
        });

        const formatRef = {};
        assert.strictEqual(sheet.readBlank(row, 0) instanceof xl.Format, true);
        assert.strictEqual(sheet.readBlank(row, 0, formatRef) instanceof xl.Format, true);
        assert.strictEqual(formatRef.format instanceof xl.Format, true);
        assert.throws(() => {
            sheet.readBlank(row, 1);
        });

        row++;
    });

    it('sheet.writeBlank writes a blank cell', () => {
        assert.throws(() => {
            sheet.writeBlank();
        });
        assert.throws(() => {
            sheet.writeBlank.call({}, row, 0, format);
        });

        assert.throws(() => {
            sheet.writeBlank(row, 0, wrongFormat);
        });
        assert.throws(() => {
            sheet.writeBlank(row, 0);
        });
        assert.strictEqual(sheet.writeBlank(row, 0, format), sheet);

        row++;
    });

    it('sheet.readFormula and sheet.writeFormula read and write a formula', () => {
        assert.throws(() => sheet.writeFormula.call(sheet, row, 'a', '=SUM(A1:A10)'));
        assert.throws(() => sheet.writeFormula.call({}, row, 0, '=SUM(A1:A10)'));

        sheet.writeFormula(row, 0, '=SUM(A1:A10)');
        sheet.writeNum(row, 1, 10);

        assert.throws(() => {
            sheet.readFormula();
        });
        assert.throws(() => {
            sheet.readFormula.call({}, row, 0);
        });

        const formatRef = {};
        assert.throws(() => {
            sheet.readFormula(row, 1);
        });
        assert.strictEqual(sheet.readFormula(row, 0), 'SUM(A1:A10)');
        assert.strictEqual(sheet.readFormula(row, 0, formatRef), 'SUM(A1:A10)');
        assert.strictEqual(formatRef.format instanceof xl.Format, true);

        row++;
    });

    it('sheet.writeFormulaNum write a formula and a numeric value', () => {
        assert.throws(() => sheet.writeFormulaNum.call(sheet, row, 'a', '=SUM(A1:A10)', 1));
        assert.throws(() => sheet.writeFormulaNum.call({}, row, 0, '=SUM(A1:A10)', 1));

        sheet.writeFormulaNum(row, 0, '=SUM(A1:A10)', 1);

        assert.strictEqual(sheet.readFormula(row, 0), 'SUM(A1:A10)');

        row++;
    });

    it('sheet.writeFormulaStr write a formula and a numeric value', () => {
        assert.throws(() => sheet.writeFormulaStr.call(sheet, row, 'a', '=SUM(A1:A10)', 'a'));
        assert.throws(() => sheet.writeFormulaStr.call({}, row, 0, '=SUM(A1:A10)', 'a'));

        sheet.writeFormulaStr(row, 0, '=SUM(A1:A10)', 'a');

        assert.strictEqual(sheet.readFormula(row, 0), 'SUM(A1:A10)');

        row++;
    });

    it('sheet.writeFormulaBool write a formula and a numeric value', () => {
        assert.throws(() => sheet.writeFormulaBool.call(sheet, row, 'a', '=A1>A2', false));
        assert.throws(() => sheet.writeFormulaBool.call({}, row, 0, '=A1>A2', false));

        sheet.writeFormulaBool(row, 0, '=A1>A2', false);

        assert.strictEqual(sheet.readFormula(row, 0), 'A1>A2');

        row++;
    });

    it('sheet.writeComment writes a comment', () => {
        assert.throws(() => {
            sheet.writeComment();
        });
        assert.throws(() => {
            sheet.writeComment.call({}, row, 0, 'comment');
        });

        assert.strictEqual(sheet.writeComment(row, 0, 'comment'), sheet);

        row++;
    });

    it('sheet.removeComment removes a comment', () => {
        sheet.writeComment(row, 0, 'comment');

        assert.throws(() => sheet.removeComment.call(sheet, row, '1'));
        assert.throws(() => sheet.removeComment.call({}, row, 0));

        sheet.removeComment(row, 0);

        assert.throws(() => sheet.readComment.call(sheet, row, 0));

        row++;
    });

    it('sheet.readComment reads a comment', () => {
        sheet.writeString(row, 0, 'foo');
        sheet.writeComment(row, 0, 'comment');

        assert.throws(() => {
            sheet.readComment();
        });
        assert.throws(() => {
            sheet.readComment.call({}, row, 0);
        });

        const formatRef = {};
        assert.throws(() => {
            sheet.readComment(row, 1);
        });
        assert.strictEqual(sheet.readComment(row, 0), 'comment');

        row++;
    });

    it('sheet.isDate checks whether a cell contains a date', () => {
        sheet.writeNum(row, 0, book.datePack(1980, 8, 19), book.addFormat().setNumFormat(xl.NUMFORMAT_DATE));

        assert.throws(() => sheet.isDate.call(sheet, row, 'a'));
        assert.throws(() => sheet.isDate.call({}, row, 0));

        assert.strictEqual(sheet.isDate(row - 1, 0), false);
        assert.strictEqual(sheet.isDate(row, 0), true);

        row++;
    });

    it('sheet.isRichStr checks whether a cell contains a richt string', () => {
        assert.throws(() => sheet.isRichStr.call(sheet, row, 'a'));
        assert.throws(() => sheet.isRichStr.call({}, row, 0));

        // TODO: add a positive test once rich strings are supported
        assert.strictEqual(sheet.isRichStr(row, 0), false);
    });

    it('sheet.readError reads an error', () => {
        sheet.writeStr(row, 0, '');

        assert.throws(() => {
            sheet.readError();
        });
        assert.throws(() => {
            sheet.readError.call({}, row, 0);
        });

        assert.strictEqual(sheet.readError(row, 0), xl.ERRORTYPE_NOERROR);

        row++;
    });

    it('sheet.writeError writes an error', () => {
        assert.throws(() => sheet.writeError.call(sheet, row, '0', xl.ERRORTYPE_DIV_0));
        assert.throws(() => sheet.writeError.call({}, row, 0, xl.ERRORTYPE_DIV_0));

        sheet.writeError(row, 0, xl.ERRORTYPE_DIV_0);

        assert.strictEqual(sheet.readError(row, 0), xl.ERRORTYPE_DIV_0);

        row++;
    });

    it('sheet.writeError accepts a format', () => {
        const format = book.addFormat();

        assert.throws(() => sheet.writeError.call(sheet, row, 0, xl.ERRORTYPE_DIV_0, 10));
        assert.throws(() => sheet.writeError.call({}, row, 0, xl.ERRORTYPE_DIV_0, format));

        sheet.writeError(row, 0, xl.ERRORTYPE_DIV_0, format);

        assert.strictEqual(sheet.readError(row, 0), xl.ERRORTYPE_DIV_0);

        row++;
    });

    it('sheet.colWidth reads colum width', () => {
        sheet.setCol(0, 0, 42);

        assert.throws(() => {
            sheet.colWidth();
        });
        assert.throws(() => {
            sheet.colWidth.call({}, 0);
        });

        assert.deepStrictEqual(sheet.colWidth(0), 42);
    });

    it('sheet.rowHeight reads row Height', () => {
        sheet.setRow(1, 42);

        assert.throws(() => {
            sheet.rowHeight();
        });
        assert.throws(() => {
            sheet.rowHeight.call({}, 1);
        });

        assert.deepStrictEqual(sheet.rowHeight(1), 42);
    });

    it('sheet.colWidthPx reads column width in pixels', () => {
        assert.throws(() => sheet.colWidthPx.call(sheet, '1'));
        assert.throws(() => sheet.colWidthPx.call({}, 1));

        assert.strictEqual(typeof sheet.colWidthPx(1), 'number');
    });

    it('sheet.colWidthPx reads column width in pixels', () => {
        assert.throws(() => sheet.rowHeightPx.call(sheet, '1'));
        assert.throws(() => sheet.rowHeightPx.call({}, 1));

        assert.strictEqual(typeof sheet.rowHeightPx(1), 'number');
    });

    it('sheet.setCol configures columns', () => {
        assert.throws(() => {
            sheet.setCol();
        });
        assert.throws(() => {
            sheet.setCol.call({}, 0, 0, 42);
        });

        assert.throws(() => {
            sheet.setCol(0, 0, 42, wrongFormat);
        });

        assert.strictEqual(sheet.setCol(0, 0, 42), sheet);
        assert.strictEqual(sheet.setCol(0, 0, 42, format, false), sheet);
    });

    it('sheet.setColPx configures columms in pixels', () => {
        assert.throws(() => sheet.setColPx.call(sheet, 0, 0, '10'));
        assert.throws(() => sheet.setColPx.call({}, 0, 0, 10));

        sheet.setColPx(0, 0, 10);

        assert.strictEqual(sheet.colWidthPx(0), 10);
    });

    it('sheet.setColPx accepts format and hidden flag', () => {
        const format = book.addFormat();

        assert.throws(() => sheet.setColPx.call(sheet, 0, 0, 10, format, '1'));

        sheet.setColPx(0, 0, 12, format, false);

        assert.strictEqual(sheet.colWidthPx(0), 12);
    });

    it('sheet.setRowPx configures rows in pixels', () => {
        assert.throws(() => sheet.setRowPx.call(sheet, 1, '10'));
        assert.throws(() => sheet.setRowPx.call({}, 1, 10));

        sheet.setRowPx(1, 20);

        assert.strictEqual(sheet.rowHeightPx(1), 20);
    });

    it('sheet.setRowPx accepts format and hidden flag', () => {
        const format = book.addFormat();

        assert.throws(() => sheet.setRowPx.call(sheet, 1, 10, format, '1'));

        sheet.setRowPx(1, 12, format, false);

        assert.strictEqual(sheet.rowHeightPx(1), 12);
    });

    it('sheet.setColPx accepts format and hidden flag', () => {
        const format = book.addFormat();

        assert.throws(() => sheet.setColPx.call({}, 0, 0, 10, format, '1'));

        sheet.setColPx(0, 0, 12, format, false);

        assert.strictEqual(sheet.colWidthPx(0), 12);
    });

    it('sheet.setRow configures rows', () => {
        assert.throws(() => {
            sheet.setRow();
        });
        assert.throws(() => {
            sheet.setRow.call({}, 1, 42);
        });

        assert.throws(() => {
            sheet.setRow(1, 42, wrongFormat);
        });

        assert.strictEqual(sheet.setRow(1, 42), sheet);
        assert.strictEqual(sheet.setRow(1, 42, format, false), sheet);
    });

    it('sheet.colFormat retrieves a column format', () => {
        const format = book.addFormat();

        assert.throws(() => sheet.colFormat.call(sheet, '0'));
        assert.throws(() => sheet.colFormat.call({}, 0));

        assert.strictEqual(sheet.colFormat(0) instanceof xl.Format, true);
    });

    it('sheet.rowFormat retrieves a row format', () => {
        const format = book.addFormat();

        assert.throws(() => sheet.rowFormat.call(sheet, '1'));
        assert.throws(() => sheet.rowFormat.call({}, 1));

        assert.strictEqual(sheet.rowFormat(1) instanceof xl.Format, true);
    });

    it('sheet.rowHidden checks whether a row is hidden', () => {
        assert.throws(() => {
            sheet.rowHidden();
        });
        assert.throws(() => {
            sheet.rowHidden.call({}, row);
        });

        assert.strictEqual(sheet.rowHidden(row), false);
    });

    it('sheet.setRowHidden hides or shows a row', () => {
        assert.throws(() => {
            sheet.setRowHidden();
        });
        assert.throws(() => {
            sheet.setRowHidden.call({}, row, true);
        });

        assert.strictEqual(sheet.setRowHidden(row, true), sheet);
        assert.strictEqual(sheet.rowHidden(row), true);
        assert.strictEqual(sheet.setRowHidden(row, false), sheet);
        assert.strictEqual(sheet.rowHidden(row), false);
    });

    it('sheet.colHidden checks whether a column is hidden', () => {
        assert.throws(() => {
            sheet.colHidden();
        });
        assert.throws(() => {
            sheet.colHidden.call({}, 0);
        });

        assert.strictEqual(sheet.colHidden(0), false);
    });

    it('sheet.setColHidden hides or shows a column', () => {
        assert.throws(() => {
            sheet.setColHidden();
        });
        assert.throws(() => {
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

        assert.throws(() => sheet.setDefaultRowHeight.call(sheet, 'a'));
        assert.throws(() => sheet.setDefaultRowHeight.call({}, 66.6));

        assert.strictEqual(sheet.setDefaultRowHeight(66.6), sheet);

        assert.throws(() => sheet.defaultRowHeight.call({}));

        assert.ok(Math.abs(sheet.defaultRowHeight() - 66.6) < epsilon);
    });

    it('sheet.setMerge, sheet.getMerge and sheet.delMerge manage merged cells', () => {
        assert.throws(() => sheet.getMerge.call(sheet, row, 0));
        assert.throws(() => sheet.delMerge.call(sheet, row, 0));

        assert.throws(() => sheet.setMerge.call(sheet, 'a', row + 2, 0, 3));
        assert.throws(() => sheet.setMerge.call({}, row, row + 2, 0, 3));
        assert.strictEqual(sheet.setMerge(row, row + 2, 0, 3), sheet);

        assert.throws(() => sheet.getMerge.call(sheet, 'a', 0));
        assert.throws(() => sheet.getMerge.call({}, row, 0));
        const merge = sheet.getMerge(row, 0);
        assert.strictEqual(merge.rowFirst, row);
        assert.strictEqual(merge.rowLast, row + 2);
        assert.strictEqual(merge.colFirst, 0);
        assert.strictEqual(merge.colLast, 3);

        assert.throws(() => sheet.delMerge.call(sheet, 'a', 0));
        assert.throws(() => sheet.delMerge.call({}, row, 0));
        sheet.delMerge(row, 0);

        assert.throws(() => sheet.delMerge.call(sheet, row, 0));
    });

    it('sheet.mergeSize, sheet.merge and sheet.delMergeByIndex manage merged cells by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');
        const mergeSizeStart = sheet.mergeSize();

        assert.throws(() => sheet.mergeSize.call({}));
        assert.throws(() => sheet.merge.call(sheet, 'a'));
        assert.throws(() => sheet.merge.call({}, 1));
        assert.throws(() => sheet.delMergeByIndex.call(sheet, 'a'));
        assert.throws(() => sheet.delMergeByIndex.call({}, 1));

        sheet.setMerge(row, row + 2, 0, 3);

        assert.strictEqual(sheet.mergeSize(), mergeSizeStart + 1);
        const mergeIdx = sheet.mergeSize() - 1;

        const merge = sheet.merge(mergeIdx);
        assert.strictEqual(merge.rowFirst, row);
        assert.strictEqual(merge.rowLast, row + 2);
        assert.strictEqual(merge.colFirst, 0);
        assert.strictEqual(merge.colLast, 3);

        assert.strictEqual(sheet.delMergeByIndex(mergeIdx), sheet);
        assert.strictEqual(sheet.mergeSize(), mergeSizeStart);
    });

    it('sheet.pictureSize returns the number of pictures on a sheet', () => {
        assert.throws(() => sheet.pictureSize.call({}));
        assert.strictEqual(sheet.pictureSize(), 0);
    });

    it('sheet.setPicture and sheet.getPicture add and inspect pictures', () => {
        assert.throws(() => sheet.setPicture.call(sheet, row, true));
        assert.throws(() => sheet.setPicture.call({}, row, 0, 0, 1, 0, 0));
        assert.throws(() => sheet.setPicture.call(sheet, row, 0, 0, 1, 0, 0, 'a'));

        assert.strictEqual(sheet.setPicture(row, 0, 0, 1, 0, 0), sheet);
        assert.strictEqual(sheet.setPicture(row, 0, 0, 1, 0, 0, xl.POSITION_ONLY_MOVE), sheet);

        let idx = sheet.pictureSize() - 1;
        assert.throws(() => sheet.getPicture.call(sheet, true));
        assert.throws(() => sheet.getPicture.call({}, idx));

        let pic = sheet.getPicture(idx);
        assert.strictEqual(pic.rowTop, row);
        assert.strictEqual(pic.colLeft, 0);
        assert.strictEqual(pic.width, testPictureWidth);
        assert.strictEqual(pic.height, testPictureHeight);
        assert.strictEqual(pic.hasOwnProperty('rowBottom'), true);
        assert.strictEqual(pic.hasOwnProperty('colRight'), true);

        row = pic.rowBottom + 1;
    });

    it('ssheet.getPicture can read links', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        book.addPictureAsLink(getTestPicturePath());
        sheet.setPicture(row, 0, 0, 1, 0, 0);

        let idx = sheet.pictureSize() - 1;

        let pic = sheet.getPicture(idx);
        assert.strictEqual(typeof pic.linkPath, 'string');

        row = pic.rowBottom + 1;
    });

    it('sheet.setPicture2 adds pictures by width and height instead of scale', () => {
        assert.throws(() => sheet.setPicture2.call(sheet, row, true));
        assert.throws(() => sheet.setPicture2.call({}, row, 0, 0, 100, 100, 0, 0));
        assert.throws(() => sheet.setPicture2.call(sheet, row, 0, 0, 100, 100, 0, 0, 'a'));

        assert.strictEqual(sheet.setPicture2(row, 0, 0, 100, 200, 0, 0), sheet);
        assert.strictEqual(sheet.setPicture2(row, 0, 0, 100, 200, 0, 0, xl.POSITION_ONLY_MOVE), sheet);

        let idx = sheet.pictureSize() - 1;

        let pic = sheet.getPicture(idx);
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

        assert.throws(() => sheet.removePicture.call(sheet, row, 'a'));
        assert.throws(() => sheet.removePicture.call({}, row, 0));

        assert.strictEqual(sheet.removePicture(row, 0), sheet);

        assert.strictEqual(sheet.pictureSize(), pictureSizeOld);
    });

    it('sheet.removePictureByIndex removes a picture from the sheet by index', () => {
        const pictureSizeOld = sheet.pictureSize();
        sheet.setPicture(row, 0, 0);

        assert.throws(() => sheet.removePictureByIndex.call(sheet, 'a'));
        assert.throws(() => sheet.removePictureByIndex.call({}, pictureSizeOld - 1));

        assert.strictEqual(sheet.removePictureByIndex(pictureSizeOld - 1), sheet);

        assert.strictEqual(sheet.pictureSize(), pictureSizeOld);
    });

    it(
        'sheet.getHorPageBreak, sheet.setHorPageBreak and sheet.horPageBreakSize' + 'manage horizontal page breaks',
        () => {
            assert.throws(() => sheet.getHorPageBreakSize.call({}));
            let n = sheet.getHorPageBreakSize();

            assert.throws(() => sheet.setHorPageBreak.call(sheet, 'a'));
            assert.throws(() => sheet.setHorPageBreak.call({}, row));
            assert.strictEqual(sheet.setHorPageBreak(row), sheet);
            assert.strictEqual(sheet.getHorPageBreakSize(), n + 1);

            assert.throws(() => sheet.getHorPageBreak.call(sheet, 'a'));
            assert.throws(() => sheet.getHorPageBreak.call({}, n + 1));
            assert.strictEqual(sheet.getHorPageBreak(n), row);

            sheet.setHorPageBreak(row, false);
            assert.strictEqual(sheet.getHorPageBreakSize(), n);
        },
    );

    it(
        'sheet.getVerPageBreak, sheet.setVerPageBreak and sheet.horPageBreakSize' + 'manage vertical page breaks',
        () => {
            assert.throws(() => sheet.getVerPageBreakSize.call({}));
            let n = sheet.getVerPageBreakSize();

            assert.throws(() => sheet.setVerPageBreak.call(sheet, 'a'));
            assert.throws(() => sheet.setVerPageBreak.call({}, row));
            assert.strictEqual(sheet.setVerPageBreak(10), sheet);
            assert.strictEqual(sheet.getVerPageBreakSize(), n + 1);

            assert.throws(() => sheet.getVerPageBreak.call(sheet, 'a'));
            assert.throws(() => sheet.getVerPageBreak.call({}, n + 1));
            assert.strictEqual(sheet.getVerPageBreak(n), 10);

            sheet.setVerPageBreak(10, false);
            assert.strictEqual(sheet.getVerPageBreakSize(), n);
        },
    );

    it('sheet.split splits a sheet', () => {
        let book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo');

        assert.throws(() => sheet.split.call(sheet, 'a', 1));
        assert.throws(() => sheet.split.call({}, 2, 2));
        assert.strictEqual(sheet.split(2, 2), sheet);
    });

    it('sheet.splitInfo obtains info about a sheet split', () => {
        let book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo');

        sheet.split(2, 3);

        assert.throws(() => sheet.splitInfo.call({}));

        const result = sheet.splitInfo();

        assert.strictEqual(result.row, 2);
        assert.strictEqual(result.col, 3);
    });

    it('sheet.groupRows groups rows', () => {
        let sheet = newSheet();

        assert.throws(() => sheet.groupRows.call(sheet, 1, 'a'));
        assert.throws(() => sheet.groupRows.call({}, 1, 2));
        assert.strictEqual(sheet.groupRows(1, 2), sheet);
    });

    it('sheet.groupCols groups columns', () => {
        let sheet = newSheet();

        assert.throws(() => sheet.groupCols.call(sheet, 1, 'a'));
        assert.throws(() => sheet.groupCols.call({}, 1, 2));
        assert.strictEqual(sheet.groupCols(1, 2), sheet);
    });

    it('sheet.setGroupSummaryBelow / sheet.groupSummaryBelow manage vert summary position', () => {
        assert.throws(() => sheet.setGroupSummaryBelow.call(sheet, 1));
        assert.throws(() => sheet.setGroupSummaryBelow.call({}, true));
        assert.strictEqual(sheet.setGroupSummaryBelow(true), sheet);

        assert.throws(() => sheet.groupSummaryBelow.call({}));
        assert.strictEqual(sheet.groupSummaryBelow(), true);

        assert.strictEqual(sheet.setGroupSummaryBelow(false).groupSummaryBelow(), false);
    });

    it('sheet.setGroupSummaryRight / sheet.groupSummaryRight manage hor summary position', () => {
        assert.throws(() => sheet.setGroupSummaryRight.call(sheet, 1));
        assert.throws(() => sheet.setGroupSummaryRight.call({}, true));
        assert.strictEqual(sheet.setGroupSummaryRight(true), sheet);

        assert.throws(() => sheet.groupSummaryRight.call({}));
        assert.strictEqual(sheet.groupSummaryRight(), true);

        assert.strictEqual(sheet.setGroupSummaryRight(false).groupSummaryRight(), false);
    });

    it('sheet.clear clears the sheet', () => {
        let sheet = newSheet();

        sheet.writeStr(1, 1, 'foo').writeStr(2, 1, 'bar');

        assert.throws(() => sheet.clear.call(sheet, 'a'));
        assert.throws(() => sheet.clear.call({}, 1, 1, 1, 1));
        assert.strictEqual(sheet.clear(1, 1, 1, 1), sheet);

        assert.throws(() => {
            sheet.readStr(1, 1);
        });
        assert.strictEqual(sheet.readStr(2, 1), 'bar');
    });

    it('sheet.insertRow and sheet.insertCol insert rows and cols', () => {
        let sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        assert.throws(() => sheet.insertRow.call(sheet, 'a', 2));
        assert.throws(() => sheet.insertRow.call({}, 2, 3));
        assert.throws(() => sheet.insertCol.call(sheet, 'a', 2));
        assert.throws(() => sheet.insertCol.call({}, 2, 3));

        assert.strictEqual(sheet.insertRow(2, 3), sheet);
        assert.strictEqual(sheet.insertCol(2, 3), sheet);

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.insertRow and sheet.insertCol support updateNamedRanges', () => {
        let sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        assert.throws(() => sheet.insertRow.call(sheet, 2, 3, 'a'));
        assert.throws(() => sheet.insertRow.call({}, 2, 3, true));
        assert.throws(() => sheet.insertCol.call(sheet, 2, 3, 'a'));
        assert.throws(() => sheet.insertCol.call({}, 2, 3, true));

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

        assert.throws(() => sheet.insertRowAsync.call(sheet, 'a', 2, () => {}));
        assert.throws(() => sheet.insertRowAsync.call({}, 2, 3, () => {}));

        const insertRowResult = util.promisify(sheet.insertRowAsync.bind(sheet))(2, 3);
        assert.throws(() => book.sheetCount.call(book));

        await insertRowResult;

        assert.throws(() => sheet.insertColAsync.call(sheet, 'a', 2, () => {}));
        assert.throws(() => sheet.insertColAsync.call({}, 2, 3, () => {}));

        const insertColResult = util.promisify(sheet.insertColAsync.bind(sheet))(2, 3);
        assert.throws(() => book.sheetCount.call(book));

        await insertColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.insertRowAsync and sheet.insertColAsync insert rows support updateNamedRanges', async () => {
        const sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 2, '12').writeStr(2, 1, '21').writeStr(2, 2, '22');

        assert.throws(() => sheet.insertRowAsync.call(sheet, 2, 3, 'a', () => {}));
        assert.throws(() => sheet.insertRowAsync.call({}, 2, 3, true, () => {}));

        const insertRowResult = util.promisify(sheet.insertRowAsync.bind(sheet))(2, 3, true);
        assert.throws(() => book.sheetCount.call(book));

        await insertRowResult;

        assert.throws(() => sheet.insertColAsync.call(sheet, 2, 3, 'a', () => {}));
        assert.throws(() => sheet.insertColAsync.call({}, 2, 3, true, () => {}));

        const insertColResult = util.promisify(sheet.insertColAsync.bind(sheet))(2, 3, true);
        assert.throws(() => book.sheetCount.call(book));

        await insertColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 4), '12');
        assert.strictEqual(sheet.readStr(4, 1), '21');
        assert.strictEqual(sheet.readStr(4, 4), '22');
    });

    it('sheet.removeRow and sheet.removeCol remove rows and cols', () => {
        let sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        assert.throws(() => sheet.removeRow.call(sheet, 'a', 2));
        assert.throws(() => sheet.removeRow.call({}, 2, 3));
        assert.throws(() => sheet.removeCol.call(sheet, 'a', 2));
        assert.throws(() => sheet.removeCol.call({}, 2, 3));

        assert.strictEqual(sheet.removeRow(2, 3), sheet);
        assert.strictEqual(sheet.removeCol(2, 3), sheet);

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.removeRow and sheet.removeCol support updateNamedRanges', () => {
        let sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        assert.throws(() => sheet.removeRow.call(sheet, 2, 3, 'a'));
        assert.throws(() => sheet.removeRow.call({}, 2, 3, true));
        assert.throws(() => sheet.removeCol.call(sheet, 2, 3, 'a'));
        assert.throws(() => sheet.removeCol.call({}, 2, 3, true));

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

        assert.throws(() => sheet.removeRowAsync.call(sheet, 'a', 2, () => undefined));
        assert.throws(() => sheet.removeRowAsync.call({}, 2, 3, () => undefined));

        const removeRowResult = util.promisify(sheet.removeRowAsync.bind(sheet))(2, 3);
        assert.throws(() => book.sheetCount.call(book));

        await removeRowResult;

        assert.throws(() => sheet.removeColAsync.call(sheet, 'a', 2, () => undefined));
        assert.throws(() => sheet.removeColAsync.call({}, 2, 3, () => undefined));

        const removeColResult = util.promisify(sheet.removeColAsync.bind(sheet))(2, 3);
        assert.throws(() => book.sheetCount.call(book));

        await removeColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.removeRowAsync and sheet.removeColAsync support updateNamedRanges', async () => {
        const sheet = newSheet();

        sheet.writeStr(1, 1, '11').writeStr(1, 4, '12').writeStr(4, 1, '21').writeStr(4, 4, '22');

        assert.throws(() => sheet.removeRowAsync.call(sheet, 2, 3, 'a', () => undefined));
        assert.throws(() => sheet.removeRowAsync.call({}, 2, 3, true, () => undefined));

        const removeRowResult = util.promisify(sheet.removeRowAsync.bind(sheet))(2, 3, true);
        assert.throws(() => book.sheetCount.call(book));

        await removeRowResult;

        assert.throws(() => sheet.removeColAsync.call(sheet, 2, 3, 'a', () => undefined));
        assert.throws(() => sheet.removeColAsync.call({}, 2, 3, true, () => undefined));

        const removeColResult = util.promisify(sheet.removeColAsync.bind(sheet))(2, 3, true);
        assert.throws(() => book.sheetCount.call(book));

        await removeColResult;

        assert.strictEqual(sheet.readStr(1, 1), '11');
        assert.strictEqual(sheet.readStr(1, 2), '12');
        assert.strictEqual(sheet.readStr(2, 1), '21');
        assert.strictEqual(sheet.readStr(2, 2), '22');
    });

    it('sheet.copyCell copies a cell', () => {
        sheet.writeStr(row, 0, 'baz');

        assert.throws(() => sheet.copyCell.call(sheet, row, 0, row, 'a'));
        assert.throws(() => sheet.copyCell.call({}, row, 0, row, 1));

        assert.strictEqual(sheet.copyCell(row, 0, row, 1), sheet);
        assert.strictEqual(sheet.readStr(row, 1), 'baz');

        row++;
    });

    it('sheet.firstRow, sheet.firstCol, sheet.lastRow, sheet.lastCol return ' + 'the spreadsheet limits', () => {
        let sheet = newSheet();

        sheet.writeNum(2, 1, 1).writeNum(5, 5, 1);

        assert.throws(() => sheet.firstRow.call({}));
        assert.throws(() => sheet.firstCol.call({}));
        assert.throws(() => sheet.lastRow.call({}));
        assert.throws(() => sheet.lastCol.call({}));

        assert.strictEqual(sheet.firstRow(), 2);
        assert.strictEqual(sheet.firstCol(), 1);
        assert.strictEqual(sheet.lastRow(), 6);
        assert.strictEqual(sheet.lastCol(), 6);
    });

    it(
        'sheet.firstFilledRow, sheet.firstFilledCol, sheet.lastFilledRow, sheet.lastFilledCol return ' +
            'the spreadsheet limits',
        () => {
            let sheet = newSheet();

            sheet.writeNum(2, 1, 1).writeNum(5, 5, 1);

            assert.throws(() => sheet.firstFilledRow.call({}));
            assert.throws(() => sheet.firstFilledCol.call({}));
            assert.throws(() => sheet.lastFilledRow.call({}));
            assert.throws(() => sheet.lastFilledCol.call({}));

            assert.strictEqual(sheet.firstFilledRow(), 0);
            assert.strictEqual(sheet.firstFilledCol(), 0);
            assert.strictEqual(sheet.lastFilledRow(), 6);
            assert.strictEqual(sheet.lastFilledCol(), 6);
        },
    );

    it('sheet.displayGridlines and sheet.setDisplayGridlines manage display of gridlines', () => {
        assert.throws(() => sheet.setDisplayGridlines.call(sheet, 1));
        assert.throws(() => sheet.setDisplayGridlines.call({}));
        assert.strictEqual(sheet.setDisplayGridlines(), sheet);

        assert.throws(() => sheet.displayGridlines.call({}));
        assert.strictEqual(sheet.displayGridlines(), true);
        assert.strictEqual(sheet.setDisplayGridlines(false).displayGridlines(), false);
    });

    it('sheet.printGridlines and sheet.setPrintGridlines manage print of gridlines', () => {
        assert.throws(() => sheet.setPrintGridlines.call(sheet, 1));
        assert.throws(() => sheet.setPrintGridlines.call({}));
        assert.strictEqual(sheet.setPrintGridlines(), sheet);

        assert.throws(() => sheet.printGridlines.call({}));
        assert.strictEqual(sheet.printGridlines(), true);
        assert.strictEqual(sheet.setPrintGridlines(false).printGridlines(), false);
    });

    it('sheet.zoom and sheet.setZoom control sheet zoom', () => {
        assert.throws(() => sheet.setZoom.call(sheet, true));
        assert.throws(() => sheet.setZoom.call({}, 10));
        assert.strictEqual(sheet.setZoom(10), sheet);

        assert.throws(() => sheet.zoom.call({}));
        assert.strictEqual(sheet.zoom(), 10);
        assert.strictEqual(sheet.setZoom(1).zoom(), 1);
    });

    it('sheet.setLandscape and sheet.landscape control landscape mode', () => {
        assert.throws(() => sheet.setLandscape.call(sheet, 1));
        assert.throws(() => sheet.setLandscape.call({}));
        assert.strictEqual(sheet.setLandscape(), sheet);

        assert.throws(() => sheet.landscape.call({}));
        assert.strictEqual(sheet.landscape(), true);
        assert.strictEqual(sheet.setLandscape(false).landscape(), false);
    });

    it('sheet.paper and sheet.setPaper control paper format', () => {
        assert.throws(() => sheet.setPaper.call(sheet, true));
        assert.throws(() => sheet.setPaper.call({}, xl.PAPER_NOTE));
        assert.strictEqual(sheet.setPaper(xl.PAPER_NOTE), sheet);

        assert.throws(() => sheet.paper.call({}));
        assert.strictEqual(sheet.paper(), xl.PAPER_NOTE);
        assert.strictEqual(sheet.setPaper().paper(), xl.PAPER_DEFAULT);
    });

    it('sheet.header, sheet.setHeader and sheet.headerMargin control the header', () => {
        assert.throws(() => sheet.setHeader.call(sheet, 1));
        assert.throws(() => sheet.setHeader.call({}, 'fuppe'));
        assert.strictEqual(sheet.setHeader('fuppe'), sheet);

        assert.throws(() => sheet.header.call({}));
        assert.strictEqual(sheet.header(), 'fuppe');

        assert.throws(() => sheet.headerMargin.call({}));
        assert.strictEqual(sheet.headerMargin(), 0.5);
        assert.strictEqual(sheet.setHeader('fuppe', 11.11).headerMargin(), 11.11);
    });

    it('sheet.footer, sheet.setFooter and sheet.footerMargin control the footer', () => {
        assert.throws(() => sheet.setFooter.call(sheet, 1));
        assert.throws(() => sheet.setFooter.call({}, 'foppe'));
        assert.strictEqual(sheet.setFooter('foppe'), sheet);

        assert.throws(() => sheet.footer.call({}));
        assert.strictEqual(sheet.footer(), 'foppe');

        assert.throws(() => sheet.footerMargin.call({}));
        assert.strictEqual(sheet.footerMargin(), 0.5);
        assert.strictEqual(sheet.setFooter('foppe', 11.11).footerMargin(), 11.11);
    });

    ['Left', 'Right', 'Bottom', 'Top'].forEach((side) => {
        let getter = 'margin' + side,
            setter = 'setMargin' + side;

        it(util.format('sheet.%s and sheet.%s control %s margin', getter, setter, side.toLowerCase()), () => {
            assert.throws(() => sheet[setter].call(sheet, true));
            assert.throws(() => sheet[setter].call({}, 11.11));
            assert.strictEqual(sheet[setter](11.11), sheet);

            assert.throws(() => sheet[getter].call({}));
            assert.strictEqual(sheet[getter](), 11.11);
            assert.strictEqual(sheet[setter](12.12)[getter](), 12.12);
        });
    });

    it('sheet.printRowCol and sheet.setPrintRowCol control printing of row / col headers', () => {
        assert.throws(() => sheet.setPrintRowCol.call(sheet, 1));
        assert.throws(() => sheet.setPrintRowCol.call({}));
        assert.strictEqual(sheet.setPrintRowCol(), sheet);

        assert.throws(() => sheet.printRowCol.call({}));
        assert.strictEqual(sheet.printRowCol(), true);
        assert.strictEqual(sheet.setPrintRowCol(false).printRowCol(), false);
    });

    it(
        'sheet.setPrintRepeatRows / sheet.printRepeatRows, sheet.setPrintRepeatCols / sheet.printRepeatCols, sheet.clearPrintRepeat ' +
            ' control row / col repeat in print',
        () => {
            const book = new xl.Book(xl.BOOK_TYPE_XLS);
            const sheet = book.addSheet('foo');
            sheet.writeStr(1, 1, 'foo');
            sheet.writeStr(2, 2, 'foo');

            assert.throws(() => sheet.setPrintRepeatRows.call(sheet, true, 1));
            assert.throws(() => sheet.setPrintRepeatRows.call({}, 1, 2));
            assert.strictEqual(sheet.setPrintRepeatRows(1, 2), sheet);

            assert.throws(() => sheet.printRepeatRows.call({}));
            const printRepeatRows = sheet.printRepeatRows();

            assert.strictEqual(printRepeatRows.rowFirst, 1);
            assert.strictEqual(printRepeatRows.rowLast, 2);

            assert.throws(() => sheet.setPrintRepeatCols.call(sheet, true, 1));
            assert.throws(() => sheet.setPrintRepeatCols.call({}, 1, 2));
            assert.strictEqual(sheet.setPrintRepeatCols(1, 2), sheet);

            assert.throws(() => sheet.printRepeatCols.call({}));
            const printRepeatCols = sheet.printRepeatCols();

            assert.strictEqual(printRepeatCols.colFirst, 1);
            assert.strictEqual(printRepeatCols.colLast, 2);

            assert.throws(() => sheet.clearPrintRepeats.call({}));
            assert.strictEqual(sheet.clearPrintRepeats(), sheet);

            assert.throws(() => sheet.printRepeatCols.call(sheet));
            assert.throws(() => sheet.printRepeatRows.call(sheet));
        },
    );

    it('sheet.setPrintArea and sheet.clearPrintArea manage the print area', () => {
        assert.throws(() => sheet.setPrintArea.call(sheet, true, 1, 5, 5));
        assert.throws(() => sheet.setPrintArea.call({}, 1, 1, 5, 5));
        assert.strictEqual(sheet.setPrintArea(1, 1, 5, 5), sheet);

        assert.throws(() => sheet.clearPrintArea.call({}));
        assert.strictEqual(sheet.clearPrintArea(), sheet);
    });

    it(
        'sheet.getNamedRange, sheet.setNamedRange, sheet.delNamedRange, sheet.namedRangeSize and sheet.namedRange ' +
            'manage named ranges',
        () => {
            let sheet2 = newSheet();

            const assertRange = (range, rowFirst, rowLast, colFirst, colLast, name, scope) => {
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
            };

            assert.throws(() => sheet.setNamedRange.call(sheet, 'range1', 'a', 1, 5, 5));
            assert.throws(() => sheet.setNamedRange.call({}, 'range1', 1, 1, 5, 5));
            assert.strictEqual(sheet.setNamedRange('range1', 1, 1, 5, 5, xl.SCOPE_WORKBOOK), sheet);
            assert.strictEqual(sheet2.setNamedRange('range2', 2, 2, 6, 6), sheet2);

            assert.throws(() => sheet.namedRangeSize.call({}));
            assert.strictEqual(sheet.namedRangeSize(), 1);
            assert.strictEqual(sheet2.namedRangeSize(), 1);

            let range1, range2;
            assert.throws(() => sheet.getNamedRange.call(sheet, 1));
            assert.throws(() => sheet.getNamedRange.call({}, 'range1'));
            assert.throws(() => sheet2.getNamedRange.call(sheet, 'range2', xl.SCOPE_WORKBOOK));
            assertRange(sheet.getNamedRange('range1'), 1, 1, 5, 5);
            assertRange(sheet2.getNamedRange('range2'), 2, 2, 6, 6);

            assert.throws(() => sheet.namedRange.call(sheet, true));
            assert.throws(() => sheet.namedRange.call({}, 0));
            assertRange(sheet.namedRange(0), 1, 1, 5, 5, 'range1', xl.SCOPE_WORKBOOK);
            assertRange(sheet2.namedRange(0), 2, 2, 6, 6, 'range2');

            assert.throws(() => sheet.delNamedRange.call(sheet, 1));
            assert.throws(() => sheet.delNamedRange.call({}, 'range1'));
            assert.strictEqual(sheet.delNamedRange('range1'), sheet);
            assert.strictEqual(sheet.namedRangeSize(), 0);
        },
    );

    it('sheet.tableSize gets the number of tables on the sheet', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(getXlsxTableFile());

        const sheet = book.getSheet(0);

        assert.throws(() => sheet.tableSize.call({}));

        assert.strictEqual(sheet.tableSize(), 1);
    });

    it('sheet.table retrieves a table by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(getXlsxTableFile());

        const sheet = book.getSheet(0);

        assert.throws(() => sheet.table.call(sheet, 'a'));
        assert.throws(() => sheet.table.call({}, 0));

        assert.deepStrictEqual(sheet.table(0), { rowFirst: 1, rowLast: 10, colFirst: 1, colLast: 4, totalRowCount: 1 });
    });

    it('sheet.getTable retrieves a table by nyme', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(getXlsxTableFile());

        const sheet = book.getSheet(0);

        assert.throws(() => sheet.getTable.call(sheet, 1));
        assert.throws(() => sheet.getTable.call({}, 'Table'));

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

        assert.throws(() => sheet.addTable.call({}));
        assert.throws(() => sheet.addTable.call(sheet, 'T', 0, 5, 0, 'a'));

        assert.ok(sheet.addTable('Table1', 0, 5, 0, 3) instanceof xl.Table);
    });

    it('sheet.getTableByName retrieves a table by name', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');
        sheet.addTable('MyTable', 0, 5, 0, 3);

        assert.throws(() => sheet.getTableByName.call(sheet, 1));
        assert.throws(() => sheet.getTableByName.call({}, 'MyTable'));

        assert.ok(sheet.getTableByName('MyTable') instanceof xl.Table);
    });

    it('sheet.getTableByIndex retrieves a table by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');
        sheet.addTable('MyTable', 0, 5, 0, 3);

        assert.throws(() => sheet.getTableByIndex.call(sheet, 'a'));
        assert.throws(() => sheet.getTableByIndex.call({}, 0));

        assert.ok(sheet.getTableByIndex(0) instanceof xl.Table);
    });

    it('sheet.name and sheet.setName manage the sheet name', () => {
        let sheet = book.addSheet('bazzaraz');

        assert.throws(() => sheet.name.call({}));
        assert.strictEqual(sheet.name(), 'bazzaraz');

        assert.throws(() => sheet.setName.call(sheet, 1));
        assert.throws(() => sheet.setName.call({}, 'fooinator'));
        assert.strictEqual(sheet.setName('fooinator'), sheet);
        assert.strictEqual(sheet.name(), 'fooinator');
    });

    it('sheet.protect and sheet.setProtect manage the protection flag', () => {
        assert.throws(() => sheet.setProtect.call(sheet, 1));
        assert.throws(() => sheet.setProtect.call({}));
        assert.strictEqual(sheet.setProtect(), sheet);

        assert.throws(() => sheet.protect.call({}));
        assert.strictEqual(sheet.protect(), true);
        assert.strictEqual(sheet.setProtect(false).protect(), false);
    });

    it('sheet.hidden and sheet.setHidden control wether a sheet is visible', () => {
        assert.throws(() => sheet.setHidden.call(sheet, true));
        assert.throws(() => sheet.setHidden.call({}));
        assert.strictEqual(sheet.setHidden(), sheet);

        assert.throws(() => sheet.hidden.call({}));
        assert.strictEqual(sheet.hidden(), xl.SHEETSTATE_HIDDEN);
        assert.strictEqual(sheet.setHidden(xl.SHEETSTATE_VISIBLE).hidden(), xl.SHEETSTATE_VISIBLE);
    });

    it('sheet.getTopLeftView and sheet.setTopLeftView manage top left visible edge of the sheet', () => {
        let sheet = newSheet();

        assert.throws(() => sheet.setTopLeftView.call(sheet, 5, true));
        assert.throws(() => sheet.setTopLeftView.call({}, 5, 6));
        assert.strictEqual(sheet.setTopLeftView(5, 6), sheet);

        assert.throws(() => sheet.getTopLeftView.call({}));
        let result = sheet.getTopLeftView();
        assert.strictEqual(result.row, 5);
        assert.strictEqual(result.col, 6);
    });

    it('sheet.addrToRowCol translates a string address to row / col', () => {
        assert.throws(() => sheet.addrToRowCol.call(sheet, 1));
        assert.throws(() => sheet.addrToRowCol.call({}, 'A1'));

        let result = sheet.addrToRowCol('A1');
        assert.strictEqual(result.row, 0);
        assert.strictEqual(result.col, 0);
        assert.strictEqual(result.rowRelative, true);
        assert.strictEqual(result.colRelative, true);
    });

    it('sheet.rowColToAddr builds an address string', () => {
        assert.throws(() => sheet.rowColToAddr.call(sheet, 0, 'a'));
        assert.throws(() => sheet.rowColToAddr.call({}, 0, 0));
        assert.strictEqual(sheet.rowColToAddr(0, 0), 'A1');
        assert.strictEqual(sheet.rowColToAddr(0, 0, false, false), '$A$1');
    });

    it('the hyperlink family of functions manages hyperlinks', () => {
        assert.throws(() => sheet.hyperlinkSize.call({}));
        assert.strictEqual(sheet.hyperlinkSize(), 0);

        assert.throws(() => sheet.addHyperlink.call(sheet, 1, row, row + 2, 1, 3));
        assert.throws(() => sheet.addHyperlink.call({}, 'http://foo.bar', row, 1, row + 2, 3));

        assert.strictEqual(sheet.addHyperlink('http://foo.bar', row, row + 2, 1, 3), sheet);
        assert.strictEqual(sheet.hyperlinkSize(), 1);

        assert.throws(() => sheet.hyperlinkIndex.call(sheet, 'a', 1));
        assert.throws(() => sheet.hyperlinkIndex.call({}, row, 4));

        assert.strictEqual(sheet.hyperlinkIndex(row, 4), -1);
        assert.strictEqual(sheet.hyperlinkIndex(row, 1), 1);

        assert.throws(() => sheet.hyperlink.call(sheet, 'a'));
        assert.throws(() => sheet.hyperlink.call({}, 0));

        assert.deepStrictEqual(sheet.hyperlink(0), {
            colFirst: 1,
            colLast: 3,
            rowFirst: row,
            rowLast: row + 2,
            hyperlink: 'http://foo.bar',
        });

        assert.throws(() => sheet.delHyperlink.call(sheet, 'a'));
        assert.throws(() => sheet.delHyperlink.call({}, 0));

        assert.strictEqual(sheet.delHyperlink(0), sheet);
        assert.strictEqual(sheet.hyperlinkSize(), 0);

        row++;
    });

    it('the autofilter family of functions manages the auto filter', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        assert.throws(() => sheet.isAutoFilter.call({}));

        assert.strictEqual(sheet.isAutoFilter(), false);

        assert.throws(() => sheet.autoFilter.call({}));

        assert.ok(sheet.autoFilter() instanceof xl.AutoFilter);

        assert.throws(() => sheet.applyFilter.call({}));

        assert.strictEqual(sheet.applyFilter(), sheet);
        assert.strictEqual(sheet.isAutoFilter(), true);

        assert.throws(() => sheet.removeFilter.call({}));

        assert.strictEqual(sheet.removeFilter(), sheet);
        assert.strictEqual(sheet.isAutoFilter(), false);
    });

    it('sheet.setAutoFitArea sets the auto fit area', () => {
        assert.throws(() => sheet.setAutoFitArea.call(sheet, 1, 1, 10, 'a'));

        assert.strictEqual(sheet.setAutoFitArea(1, 1, 10, 10), sheet);
    });

    it('sheet.setTabColor, sheet.getTabColor and sheet.tabColor manage the tab color', () => {
        assert.throws(() => sheet.setTabColor.call(sheet, 'a'));
        assert.throws(() => sheet.setTabColor.call({}, xl.COLOR_BLACK));
        assert.throws(() => sheet.setTabColor.call(sheet, 255, 200, 'a'));
        assert.throws(() => sheet.setTabColor.call({}, 255, 200, 180));
        assert.throws(() => sheet.tabColor.call({}));

        assert.strictEqual(sheet.setTabColor(xl.COLOR_BLACK), sheet);
        assert.strictEqual(sheet.tabColor(), xl.COLOR_BLACK);

        assert.strictEqual(sheet.setTabColor(255, 200, 180), sheet);
        assert.deepStrictEqual(sheet.getTabColor(), { red: 255, green: 200, blue: 180 });
    });

    it('sheet.addIgnoredError configures the ignored error class for a range of cells', () => {
        assert.throws(() => sheet.addIgnoredError.call(sheet, row, 1, row + 2, 3, 'a'));
        assert.throws(() => sheet.addIgnoredError.call({}, row, 1, row + 2, 3, xl.IERR_EVAL_ERROR));

        assert.strictEqual(sheet.addIgnoredError(row, 1, row + 2, 3, xl.IERR_TWODIG_TEXTYEAR), sheet);
    });

    it('sheet.addDataValidation adds a string data validation for a range of cell', () => {
        assert.throws(() =>
            sheet.addDataValidation.call(
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
            ),
        );

        assert.throws(() =>
            sheet.addDataValidation.call(
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
            ),
        );

        assert.throws(() =>
            sheet.addDataValidation.call(sheet, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 1),
        );
        assert.throws(() =>
            sheet.addDataValidation.call({}, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 'a'),
        );
        assert.throws(() =>
            sheet.addDataValidation.call(sheet, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5),
        );

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
        assert.throws(() =>
            sheet.addDataValidationDouble.call(
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
            ),
        );

        assert.throws(() =>
            sheet.addDataValidationDouble.call(
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
            ),
        );

        assert.throws(() =>
            sheet.addDataValidationDouble.call(
                sheet,
                xl.VALIDATION_TYPE_WHOLE,
                xl.VALIDATION_OP_EQUAL,
                1,
                5,
                1,
                5,
                1,
                'a',
            ),
        );
        assert.throws(() =>
            sheet.addDataValidationDouble.call({}, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 1, 2),
        );
        assert.throws(() =>
            sheet.addDataValidationDouble.call(sheet, xl.VALIDATION_TYPE_WHOLE, xl.VALIDATION_OP_EQUAL, 1, 5, 1, 5, 1),
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
        book.loadSync(getXlsmFormControlFile());

        const sheet = book.getSheet(0);

        assert.throws(() => sheet.formControlSize.call({}));

        assert.strictEqual(sheet.formControlSize(), 1);
    });

    it('sheet.formControl retrieves a form control by index', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        book.loadSync(getXlsmFormControlFile());

        const sheet = book.getSheet(0);

        assert.throws(() => sheet.formControl.call(sheet, 'a'));
        assert.throws(() => sheet.formControl.call({}, 0));

        assert.ok(sheet.formControl(0) instanceof xl.FormControl);
    });

    it('sheet.addConditionalFormatting adds a conditional formatting rule', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        assert.throws(() => sheet.addConditionalFormatting.call({}));
        assert.throws(() => sheet.addConditionalFormatting.call(sheet, 0, 0, 1, 'a'));

        assert.ok(sheet.addConditionalFormatting(0, 0, 1, 1) instanceof xl.ConditionalFormatting);
    });

    it('sheet.getActiveSell and sheet.setActiveCell manage the active cell', () => {
        assert.throws(() => sheet.setActiveCell.call(sheet, 1, 'a'));
        assert.throws(() => sheet.setActiveCell.call({}, 1, 1));

        assert.strictEqual(sheet.setActiveCell(1, 1), sheet);

        assert.throws(() => sheet.getActiveCell.call({}));

        assert.deepStrictEqual(sheet.getActiveCell(), { row: 1, col: 1 });
    });

    it('sheet.addSectionRange, sheet.removeSelection and sheet.selectionRange manage the selection range', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        assert.throws(() => sheet.removeSelection.call({}));
        assert.strictEqual(sheet.removeSelection(), sheet);

        assert.throws(() => sheet.addSelectionRange.call(sheet, 1));
        assert.throws(() => sheet.addSelectionRange.call({}, 'A1:A12'));

        assert.strictEqual(sheet.addSelectionRange('A2:A10'), sheet);

        assert.throws(() => sheet.selectionRange.call({}));

        assert.strictEqual(sheet.selectionRange(), 'A2:A10');
    });

    it('sheet.setBorder draws borders in an area of cells', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        assert.throws(() => sheet.setBorder.call(sheet, 0, 0, 0, 0, 'a', 0));
        assert.throws(() => sheet.setBorder.call({}, 0, 0, 0, 0, xl.BORDERSTYLE_THIN, xl.COLOR_BLACK));

        assert.strictEqual(sheet.setBorder(0, 5, 0, 5, xl.BORDERSTYLE_THIN, xl.COLOR_BLACK), sheet);
    });

    it('sheet.applyFilter2 applies a filter using an AutoFilter object', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        const af = sheet.autoFilter();

        assert.throws(() => sheet.applyFilter2.call(sheet, 1));
        assert.throws(() => sheet.applyFilter2.call({}, af));

        assert.strictEqual(sheet.applyFilter2(af), sheet);
    });

    it('sheet.conditionalFormattingSize, sheet.conditionalFormatting and sheet.removeConditionalFormatting manage conditional formatting', () => {
        const book = new xl.Book(xl.BOOK_TYPE_XLSX);
        const sheet = book.addSheet('foo');

        assert.throws(() => sheet.conditionalFormattingSize.call({}));
        assert.strictEqual(sheet.conditionalFormattingSize(), 0);

        sheet.addConditionalFormatting(0, 0, 1, 1);
        assert.strictEqual(sheet.conditionalFormattingSize(), 1);

        assert.throws(() => sheet.conditionalFormatting.call(sheet, 'a'));
        assert.throws(() => sheet.conditionalFormatting.call({}, 0));

        assert.ok(sheet.conditionalFormatting(0) instanceof xl.ConditionalFormatting);

        assert.throws(() => sheet.removeConditionalFormatting.call(sheet, 'a'));
        assert.throws(() => sheet.removeConditionalFormatting.call({}, 0));

        assert.strictEqual(sheet.removeConditionalFormatting(0), sheet);
        assert.strictEqual(sheet.conditionalFormattingSize(), 0);
    });
});
