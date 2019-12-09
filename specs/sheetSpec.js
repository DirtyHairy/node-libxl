var xl = require('../lib/libxl'),
    util = require('util'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('The sheet class', function() {

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
        return book.addSheet('foo' + (sheetNameIdx++));
    }

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

    it('sheet.isDate checks whether a cell contains a date', function() {
        sheet.writeNum(row, 0, book.datePack(1980, 8, 19),
            book.addFormat().setNumFormat(xl.NUMFORMAT_DATE));

        shouldThrow(sheet.isDate, sheet, row, 'a');
        shouldThrow(sheet.isDate, {}, row, 0);

        expect(sheet.isDate(row - 1, 0)).toBe(false);
        expect(sheet.isDate(row, 0)).toBe(true);
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

    it('sheet.setMerge, sheet.getMerge and sheet.delMerge manage merged cells', function() {
        shouldThrow(sheet.getMerge, sheet, row, 0);
        shouldThrow(sheet.delMerge, sheet, row, 0);

        shouldThrow(sheet.setMerge, sheet, 'a', row + 2, 0, 3);
        shouldThrow(sheet.setMerge, {}, row, row + 2, 0, 3);
        expect(sheet.setMerge(row, row + 2, 0, 3)).toBe(sheet);

        shouldThrow(sheet.getMerge, sheet, 'a', 0);
        shouldThrow(sheet.getMerge, {}, row, 0);
        merge = sheet.getMerge(row, 0);
        expect(merge.rowFirst).toBe(row);
        expect(merge.rowLast).toBe(row + 2);
        expect(merge.colFirst).toBe(0);
        expect(merge.colLast).toBe(3);

        shouldThrow(sheet.delMerge, sheet, 'a', 0);
        shouldThrow(sheet.delMerge, {}, row, 0);
        sheet.delMerge(row, 0);

        shouldThrow(sheet.delMerge, sheet, row, 0);
    });

    it('sheet.pictureSize returns the number of pictures on a sheet', function() {
        shouldThrow(sheet.pictureSize, {});
        sheet.pictureSize();
    });

    it ('sheet.setPicture and sheet.getPicture add and inspect pictures', function() {
        shouldThrow(sheet.setPicture, sheet, row, true);
        shouldThrow(sheet.setPicture, {}, row, 0, 0, 1, 0, 0);

        expect(sheet.setPicture(row, 0, 0, 1, 0, 0)).toBe(sheet);

        var idx = sheet.pictureSize() - 1;
        shouldThrow(sheet.getPicture, sheet, true);
        shouldThrow(sheet.getPicture, {}, idx);

        var pic = sheet.getPicture(idx);
        expect(pic.rowTop).toBe(row);
        expect(pic.colLeft).toBe(0);
        expect(pic.width).toBe(testUtils.testPictureWidth);
        expect(pic.height).toBe(testUtils.testPictureHeight);
        expect(pic.hasOwnProperty('rowBottom')).toBe(true);
        expect(pic.hasOwnProperty('colRight')).toBe(true);

        row = pic.rowBottom + 1;
    });

    it ('sheet.setPicture2 adds pictures by width and height instead of scale', function() {
        shouldThrow(sheet.setPicture2, sheet, row, true);
        shouldThrow(sheet.setPicture2, {}, row, 0, 0, 100, 100, 0, 0);

        expect(sheet.setPicture2(row, 0, 0, 100, 200, 0, 0)).toBe(sheet);

        var idx = sheet.pictureSize() - 1;

        var pic = sheet.getPicture(idx);
        expect(pic.rowTop).toBe(row);
        expect(pic.colLeft).toBe(0);
        expect(pic.width).toBe(100);
        expect(pic.height).toBe(200);
        expect(pic.hasOwnProperty('rowBottom')).toBe(true);
        expect(pic.hasOwnProperty('colRight')).toBe(true);
    });

    it('sheet.getHorPageBreak, sheet.setHorPageBreak and sheet.horPageBreakSize' +
        'manage horizontal page breaks', function()
    {
        shouldThrow(sheet.getHorPageBreakSize, {});
        var n = sheet.getHorPageBreakSize();

        shouldThrow(sheet.setHorPageBreak, sheet, 'a');
        shouldThrow(sheet.setHorPageBreak, {}, row);
        expect(sheet.setHorPageBreak(row)).toBe(sheet);
        expect(sheet.getHorPageBreakSize()).toBe(n + 1);

        shouldThrow(sheet.getHorPageBreak, sheet, 'a');
        shouldThrow(sheet.getHorPageBreak, {}, n + 1);
        expect(sheet.getHorPageBreak(n)).toBe(row);

        sheet.setHorPageBreak(row, false);
        expect(sheet.getHorPageBreakSize()).toBe(n);
    });

    it('sheet.getVerPageBreak, sheet.setVerPageBreak and sheet.horPageBreakSize' +
        'manage vertical page breaks', function()
    {
        shouldThrow(sheet.getVerPageBreakSize, {});
        var n = sheet.getVerPageBreakSize();

        shouldThrow(sheet.setVerPageBreak, sheet, 'a');
        shouldThrow(sheet.setVerPageBreak, {}, row);
        expect(sheet.setVerPageBreak(10)).toBe(sheet);
        expect(sheet.getVerPageBreakSize()).toBe(n + 1);

        shouldThrow(sheet.getVerPageBreak, sheet, 'a');
        shouldThrow(sheet.getVerPageBreak, {}, n + 1);
        expect(sheet.getVerPageBreak(n)).toBe(10);

        sheet.setVerPageBreak(10, false);
        expect(sheet.getVerPageBreakSize()).toBe(n);
    });

    it('sheet.split splits a sheet', function() {
        var book = new xl.Book(xl.BOOK_TYPE_XLS),
            sheet = book.addSheet('foo');

        shouldThrow(sheet.split, sheet, 'a', 1);
        shouldThrow(sheet.split, {}, 2, 2);
        expect(sheet.split(2, 2)).toBe(sheet);
    });

    it('sheet.groupRows groups rows', function() {
        var sheet = newSheet();

        shouldThrow(sheet.groupRows, sheet, 1, 'a');
        shouldThrow(sheet.groupRows, {}, 1, 2);
        expect(sheet.groupRows(1, 2)).toBe(sheet);
    });

    it('sheet.groupCols groups columns', function() {
        var sheet = newSheet();

        shouldThrow(sheet.groupCols, sheet, 1, 'a');
        shouldThrow(sheet.groupCols, {}, 1, 2);
        expect(sheet.groupCols(1, 2)).toBe(sheet);
    });

    it('sheet.setGroupSummaryBelow / sheet.groupSummaryBelow manage vert summary position', function() {
        shouldThrow(sheet.setGroupSummaryBelow, sheet, 1);
        shouldThrow(sheet.setGroupSummaryBelow, {}, true);
        expect(sheet.setGroupSummaryBelow(true)).toBe(sheet);

        shouldThrow(sheet.groupSummaryBelow, {});
        expect(sheet.groupSummaryBelow()).toBe(true);

        expect(sheet.setGroupSummaryBelow(false).groupSummaryBelow()).toBe(false);
    });

    it('sheet.setGroupSummaryRight / sheet.groupSummaryRight manage hor summary position', function() {
        shouldThrow(sheet.setGroupSummaryRight, sheet, 1);
        shouldThrow(sheet.setGroupSummaryRight, {}, true);
        expect(sheet.setGroupSummaryRight(true)).toBe(sheet);

        shouldThrow(sheet.groupSummaryRight, {});
        expect(sheet.groupSummaryRight()).toBe(true);

        expect(sheet.setGroupSummaryRight(false).groupSummaryRight()).toBe(false);
    });

    it('sheet.clear clears the sheet', function() {
        var sheet = newSheet();

        sheet
            .writeStr(1, 1, 'foo')
            .writeStr(2, 1, 'bar');

        shouldThrow(sheet.clear, sheet, 'a');
        shouldThrow(sheet.clear, {}, 1, 1, 1, 1);
        expect(sheet.clear(1, 1, 1, 1)).toBe(sheet);

        expect(function() {sheet.readStr(1, 1);}).toThrow();
        expect(sheet.readStr(2, 1)).toBe('bar');
    });

    it('sheet.insertRow and sheet.insertCol insert rows and cols', function() {
        var sheet = newSheet();

        sheet
            .writeStr(1, 1, '11')
            .writeStr(1, 2, '12')
            .writeStr(2, 1, '21')
            .writeStr(2, 2, '22');

        shouldThrow(sheet.insertRow, sheet, 'a', 2);
        shouldThrow(sheet.insertRow, {}, 2, 3);
        shouldThrow(sheet.insertCol, sheet, 'a', 2);
        shouldThrow(sheet.insertCol, {}, 2, 3);

        expect(sheet.insertRow(2, 3)).toBe(sheet);
        expect(sheet.insertCol(2, 3)).toBe(sheet);

        expect(sheet.readStr(1, 1)).toBe('11');
        expect(sheet.readStr(1, 4)).toBe('12');
        expect(sheet.readStr(4, 1)).toBe('21');
        expect(sheet.readStr(4, 4)).toBe('22');
    });

    it('sheet.insertRowAsync and sheet.insertColAsync insert rows and cols in async mode', function() {
        var sheet = newSheet(),
            done = false;

        sheet
            .writeStr(1, 1, '11')
            .writeStr(1, 2, '12')
            .writeStr(2, 1, '21')
            .writeStr(2, 2, '22');

        runs(function() {
            shouldThrow(sheet.insertRowAsync, sheet, 'a', 2, function() {});
            shouldThrow(sheet.insertRowAsync, {}, 2, 3, function() {});
            expect(sheet.insertRowAsync(2, 3, step1)).toBe(sheet);
            shouldThrow(sheet.name, sheet);

            function step1(err) {
                expect(err).toBeUndefined();

                shouldThrow(sheet.insertColAsync, sheet, 'a', 2, function() {});
                shouldThrow(sheet.insertColAsync, {}, 2, 3, function() {});
                expect(sheet.insertColAsync(2, 3, step2)).toBe(sheet);
                shouldThrow(sheet.name, sheet);
            }

            function step2(err) {
                expect(err).toBeUndefined();

                done = true;
            }
        });

        waitsFor(function() {
            return done;
        }, 3000, 'insertRowAsync and insertColAsync to terminate');

        runs(function() {
            expect(sheet.readStr(1, 1)).toBe('11');
            expect(sheet.readStr(1, 4)).toBe('12');
            expect(sheet.readStr(4, 1)).toBe('21');
            expect(sheet.readStr(4, 4)).toBe('22');
        });
    });

    it('sheet.removeRow and sheet.removeCol remove rows and cols', function() {
        var sheet = newSheet();

        sheet
            .writeStr(1, 1, '11')
            .writeStr(1, 4, '12')
            .writeStr(4, 1, '21')
            .writeStr(4, 4, '22');

        shouldThrow(sheet.removeRow, sheet, 'a', 2);
        shouldThrow(sheet.removeRow, {}, 2, 3);
        shouldThrow(sheet.removeCol, sheet, 'a', 2);
        shouldThrow(sheet.removeCol, {}, 2, 3);

        expect(sheet.removeRow(2, 3)).toBe(sheet);
        expect(sheet.removeCol(2, 3)).toBe(sheet);

        expect(sheet.readStr(1, 1)).toBe('11');
        expect(sheet.readStr(1, 2)).toBe('12');
        expect(sheet.readStr(2, 1)).toBe('21');
        expect(sheet.readStr(2, 2)).toBe('22');
    });

    it('sheet.removeRowAsync and sheet.removeColAsync remove rows and cols in async mode', function() {
        var sheet = newSheet(),
            done = false;

        sheet
            .writeStr(1, 1, '11')
            .writeStr(1, 4, '12')
            .writeStr(4, 1, '21')
            .writeStr(4, 4, '22');

        runs(function() {
            shouldThrow(sheet.removeRowAsync, sheet, 'a', 2, step1);
            shouldThrow(sheet.removeRowAsync, {}, 2, 3, step1);
            expect(sheet.removeRowAsync(2, 3, step1)).toBe(sheet);
            shouldThrow(sheet.name, sheet);

            function step1(err) {
                expect(err).toBeUndefined();

                shouldThrow(sheet.removeColAsync, sheet, 'a', 2, step2);
                shouldThrow(sheet.removeColAsync, {}, 2, 3, step2);
                expect(sheet.removeColAsync(2, 3, step2)).toBe(sheet);
                shouldThrow(sheet.name, sheet);
            }

            function step2(err) {
                expect(err).toBeUndefined();
                done = true;
            }
        });

        waitsFor(function() {
            return done;
        }, 3000, 'removeRowAsync and removeColAsync to finish');

        runs(function() {
            expect(sheet.readStr(1, 1)).toBe('11');
            expect(sheet.readStr(1, 2)).toBe('12');
            expect(sheet.readStr(2, 1)).toBe('21');
            expect(sheet.readStr(2, 2)).toBe('22');
        });
    });


    it('sheet.copyCell copies a cell', function() {
        sheet.writeStr(row, 0, 'baz');

        shouldThrow(sheet.copyCell, sheet, row, 0, row, 'a');
        shouldThrow(sheet.copyCell, {}, row, 0, row, 1);

        expect(sheet.copyCell(row, 0, row, 1)).toBe(sheet);
        expect(sheet.readStr(row, 1)).toBe('baz');

        row++;
    });

    it('sheet.firstRow, sheet.firstCol, sheet.lastRow, sheet.lastCol return ' +
        'the spreadsheet limits', function()
    {
        var sheet = newSheet();

        sheet
            .writeNum(2, 1, 1)
            .writeNum(5, 5, 1);

        shouldThrow(sheet.firstRow, {});
        shouldThrow(sheet.firstCol, {});
        shouldThrow(sheet.lastRow, {});
        shouldThrow(sheet.lastCol, {});

        expect(sheet.firstRow()).toBe(2);
        expect(sheet.firstCol()).toBe(1);
        expect(sheet.lastRow()).toBe(6);
        expect(sheet.lastCol()).toBe(6);
    });

    it('sheet.displayGridlines and sheet.setDisplayGridlines manage display of gridlines', function() {
        shouldThrow(sheet.setDisplayGridlines, sheet, 1);
        shouldThrow(sheet.setDisplayGridlines, {});
        expect(sheet.setDisplayGridlines()).toBe(sheet);

        shouldThrow(sheet.displayGridlines, {});
        expect(sheet.displayGridlines()).toBe(true);
        expect(sheet.setDisplayGridlines(false).displayGridlines()).toBe(false);
    });

    it('sheet.printGridlines and sheet.setPrintGridlines manage print of gridlines', function() {
        shouldThrow(sheet.setPrintGridlines, sheet, 1);
        shouldThrow(sheet.setPrintGridlines, {});
        expect(sheet.setPrintGridlines()).toBe(sheet);

        shouldThrow(sheet.printGridlines, {});
        expect(sheet.printGridlines()).toBe(true);
        expect(sheet.setPrintGridlines(false).printGridlines()).toBe(false);
    });

    it('sheet.zoom and sheet.setZoom control sheet zoom', function() {
        shouldThrow(sheet.setZoom, sheet, true);
        shouldThrow(sheet.setZoom, {}, 10);
        expect(sheet.setZoom(10)).toBe(sheet);

        shouldThrow(sheet.zoom, {});
        expect(sheet.zoom()).toBe(10);
        expect(sheet.setZoom(1).zoom()).toBe(1);
    });

    it('sheet.setLandscape and sheet.landscape control landscape mode', function() {
        shouldThrow(sheet.setLandscape, sheet, 1);
        shouldThrow(sheet.setLandscape, {});
        expect(sheet.setLandscape()).toBe(sheet);

        shouldThrow(sheet.landscape, {});
        expect(sheet.landscape()).toBe(true);
        expect(sheet.setLandscape(false).landscape()).toBe(false);
    });

    it('sheet.paper and sheet.setPaper control paper format', function() {
        shouldThrow(sheet.setPaper, sheet, true);
        shouldThrow(sheet.setPaper, {}, xl.PAPER_NOTE);
        expect(sheet.setPaper(xl.PAPER_NOTE)).toBe(sheet);

        shouldThrow(sheet.paper, {});
        expect(sheet.paper()).toBe(xl.PAPER_NOTE);
        expect(sheet.setPaper().paper()).toBe(xl.PAPER_DEFAULT);
    });

    it('sheet.header, sheet.setHeader and sheet.headerMargin control the header', function() {
        shouldThrow(sheet.setHeader, sheet, 1);
        shouldThrow(sheet.setHeader, {}, 'fuppe');
        expect(sheet.setHeader('fuppe')).toBe(sheet);

        shouldThrow(sheet.header, {});
        expect(sheet.header()).toBe('fuppe');

        shouldThrow(sheet.headerMargin, {});
        expect(sheet.headerMargin()).toBe(0.5);
        expect(sheet.setHeader('fuppe', 11.11).headerMargin()).toBe(11.11);
    });

    it('sheet.footer, sheet.setFooter and sheet.footerMargin control the footer', function() {
        shouldThrow(sheet.setFooter, sheet, 1);
        shouldThrow(sheet.setFooter, {}, 'foppe');
        expect(sheet.setFooter('foppe')).toBe(sheet);

        shouldThrow(sheet.footer, {});
        expect(sheet.footer()).toBe('foppe');

        shouldThrow(sheet.footerMargin, {});
        expect(sheet.footerMargin()).toBe(0.5);
        expect(sheet.setFooter('foppe', 11.11).footerMargin()).toBe(11.11);
    });

    ['Left', 'Right', 'Bottom', 'Top'].forEach(function(side) {
        var getter = 'margin' + side,
            setter = 'setMargin' + side;

        it(util.format('sheet.%s and sheet.%s control %s margin',
            getter, setter, side.toLowerCase()), function()
        {
            shouldThrow(sheet[setter], sheet, true);
            shouldThrow(sheet[setter], {}, 11.11);
            expect(sheet[setter](11.11)).toBe(sheet);

            shouldThrow(sheet[getter], {});
            expect(sheet[getter]()).toBe(11.11);
            expect(sheet[setter](12.12)[getter]()).toBe(12.12);
        });
    });

    it('sheet.printRowCol and sheet.setPrintRowCol control printing of row / col headers', function() {
        shouldThrow(sheet.setPrintRowCol, sheet, 1);
        shouldThrow(sheet.setPrintRowCol, {});
        expect(sheet.setPrintRowCol()).toBe(sheet);

        shouldThrow(sheet.printRowCol, {});
        expect(sheet.printRowCol()).toBe(true);
        expect(sheet.setPrintRowCol(false).printRowCol()).toBe(false);
    });

    it('sheet.setPrintRepeatRows, sheet.setPrintRepeatCols, sheet.clearPrintRepeat ' +
        ' control row / col repeat in print', function()
    {
        shouldThrow(sheet.setPrintRepeatRows, sheet, true, 1);
        shouldThrow(sheet.setPrintRepeatRows, {}, 1, 1);
        expect(sheet.setPrintRepeatRows(1, 1)).toBe(sheet);

        shouldThrow(sheet.setPrintRepeatCols, sheet, true, 1);
        shouldThrow(sheet.setPrintRepeatCols, {}, 1, 1);
        expect(sheet.setPrintRepeatCols(1, 1)).toBe(sheet);

        shouldThrow(sheet.clearPrintRepeats, {});
        expect(sheet.clearPrintRepeats()).toBe(sheet);
    });

    it('sheet.setPrintArea and sheet.clearPrintArea manage the print area', function() {
        shouldThrow(sheet.setPrintArea, sheet, true, 1, 5, 5);
        shouldThrow(sheet.setPrintArea, {}, 1, 1, 5, 5);
        expect(sheet.setPrintArea(1, 1, 5, 5)).toBe(sheet);

        shouldThrow(sheet.clearPrintArea, {});
        expect(sheet.clearPrintArea()).toBe(sheet);
    });

    it('sheet.getNamedRange, sheet.setNamedRange, sheet.delNamedRange, sheet.namedRangeSize and sheet.namedRange ' +
        'manage named ranges', function()
    {
        var sheet2 = newSheet();

        function assertRange(range, rowFirst, rowLast, colFirst, colLast, name, scope) {
            expect(range.rowFirst).toBe(rowFirst);
            expect(range.rowLast).toBe(rowLast);
            expect(range.colFirst).toBe(colFirst);
            expect(range.colLast).toBe(colLast);
            if (typeof(name) !== 'undefined') {
                expect(range.name).toBe(name);
            }
            if (typeof(scope) !== 'undefined') {
                expect(range.scopeId).toBe(scope);
            }
        }

        shouldThrow(sheet.setNamedRange, sheet, 'range1', 'a', 1, 5, 5);
        shouldThrow(sheet.setNamedRange, {}, 'range1', 1, 1, 5, 5);
        expect(sheet.setNamedRange('range1', 1, 1, 5, 5, xl.SCOPE_WORKBOOK)).toBe(sheet);
        expect(sheet2.setNamedRange('range2', 2, 2, 6, 6)).toBe(sheet2);

        shouldThrow(sheet.namedRangeSize, {});
        expect(sheet.namedRangeSize()).toBe(1);
        expect(sheet2.namedRangeSize()).toBe(1);

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
        expect(sheet.delNamedRange('range1')).toBe(sheet);
        expect(sheet.namedRangeSize()).toBe(0);
    });

    it('sheet.name and sheet.setName manage the sheet name', function() {
        var sheet = book.addSheet('bazzaraz');

        shouldThrow(sheet.name, {});
        expect(sheet.name()).toBe('bazzaraz');

        shouldThrow(sheet.setName, sheet, 1);
        shouldThrow(sheet.setName, {}, 'fooinator');
        expect(sheet.setName('fooinator')).toBe(sheet);
        expect(sheet.name()).toBe('fooinator');
    });

    it('sheet.protect and sheet.setProtect manage the protection flag', function() {
        shouldThrow(sheet.setProtect, sheet, 1);
        shouldThrow(sheet.setProtect, {});
        expect(sheet.setProtect()).toBe(sheet);

        shouldThrow(sheet.protect, {});
        expect(sheet.protect()).toBe(true);
        expect(sheet.setProtect(false).protect()).toBe(false);
    });

    it('sheet.hidden and sheet.setHidden control wether a sheet is visible', function() {
        shouldThrow(sheet.setHidden, sheet, true);
        shouldThrow(sheet.setHidden, {});
        expect(sheet.setHidden()).toBe(sheet);

        shouldThrow(sheet.hidden, {});
        expect(sheet.hidden()).toBe(xl.SHEETSTATE_HIDDEN);
        expect(sheet.setHidden(xl.SHEETSTATE_VISIBLE).hidden()).toBe(xl.SHEETSTATE_VISIBLE);
    });

    it('sheet.getTopLeftView and sheet.setTopLeftView manage top left visible edge of the sheet',
        function()
    {
        var sheet = newSheet();

        shouldThrow(sheet.setTopLeftView, sheet, 5, true);
        shouldThrow(sheet.setTopLeftView, {}, 5, 6);
        expect(sheet.setTopLeftView(5, 6)).toBe(sheet);

        shouldThrow(sheet.getTopLeftView, {});
        var result = sheet.getTopLeftView();
        expect(result.row).toBe(5);
        expect(result.col).toBe(6);
    });

    it('sheet.addrToRowCol translates a string address to row / col', function() {
        shouldThrow(sheet.addrToRowCol, sheet, 1);
        shouldThrow(sheet.addrToRowCol, {}, 'A1');

        var result = sheet.addrToRowCol('A1');
        expect(result.row).toBe(0);
        expect(result.col).toBe(0);
        expect(result.rowRelative).toBe(true);
        expect(result.colRelative).toBe(true);
    });

    it('sheet.rowColToAddr builds an address string', function() {
        shouldThrow(sheet.rowColToAddr, sheet, 0, 'a');
        shouldThrow(sheet.rowColToAddr, {}, 0, 0);
        expect(sheet.rowColToAddr(0, 0)).toBe('A1');
        expect(sheet.rowColToAddr(0, 0, false, false)).toBe('$A$1');
    });
});
