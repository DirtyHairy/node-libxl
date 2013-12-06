var xl = require('./lib/libxl');

function fillSheet(sheet) {
    var currentRow = 1;

    sheet
        .writeString(currentRow++, 0, 'Some string')
        .writeString(currentRow++, 0, 'Unicode - فارسی - Қазақша');
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
