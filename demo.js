var xl = require('./lib/libxl');

function fillSheet(sheet) {
    var book = sheet.book,
        row = 1,
        format;

    sheet
        .writeStr(row, 0, 'Some string')
        .writeStr(row, 0, 'Unicode - فارسی - Қазақша');
    row++;

    format = book.addFormat();
    format.setFont(book.addFont().setSize(20));
    sheet
        .writeStr(row, 0, 'green', format);
    sheet.cellFormat(row, 0)
        .setFillPattern(xl.FILLPATTERN_SOLID)
        .setPatternForegroundColor(xl.COLOR_GREEN);
    sheet.setCellFormat(row, 1, format);

    format = {};
    sheet.readStr(row, 0, format);
    sheet.setCellFormat(row, 2, format.format);

    row++;

    sheet
        .writeStr(row, 0, 'Ten')
        .writeNum(row, 1, 10);
    row++;

    sheet
        .writeStr(row, 0, 'True')
        .writeBool(row, 1, true);
    row++;

    sheet
        .writeStr(row, 0, 'a blank cell')
        .writeStr(row, 1, 'foo')
        .writeBlank(row, 1, sheet.cellFormat(row, 0));
    row++;

    sheet
        .writeString(row, 0, 'Comment')
        .writeComment(row, 0, 'This a comment', 'from me', 200, 200);
    row++;

    sheet.writeNum(row, 0, book.datePack(1980, 8, 19),
        book.addFormat().setNumFormat(xl.NUMFORMAT_DATE));
    row++;

    sheet.split(2, 5);
}

function fillBook(book) {
    var sheet = book.addSheet('Sheet 1');

    fillSheet(sheet);
}

var xlsBook =  new xl.Book(xl.BOOK_TYPE_XLS),
    xlsxBook = new xl.Book(xl.BOOK_TYPE_XLSX);

fillBook(xlsBook);
fillBook(xlsxBook);

xlsBook.writeSync('demo.xls');
xlsxBook.writeSync('demo.xlsx');
