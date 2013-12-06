var xl = require('../lib/libxl');

describe('The sheet class', function() {

    var book = new xl.Book(xl.BOOK_TYPE_XLSX),
        sheet = book.addSheet('foo'),
        row = 1;

    it('sheet.cellType determines cell type', function() {
        sheet
            .writeString(row, 0, 'foo')
            .writeNum(row, 1, 10);

        expect(function() {sheet.cellType();}).toThrow();
        expect(function() {sheet.cellType.call({}, row, 0);}).toThrow();

        expect(sheet.cellType(row, 0)).toBe(xl.CELLTYPE_STRING);
        expect(sheet.cellType(row, 1)).toBe(xl.CELLTYPE_NUMBER);
    });

});
