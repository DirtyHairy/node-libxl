var xl = require('../lib/libxl'),
    testUtils = require('./testUtils'),
    shouldThrow = testUtils.shouldThrow;

describe('Table', () => {
    let book, sheet, table;

    beforeEach(() => {
        book = new xl.Book(xl.BOOK_TYPE_XLSX);
        sheet = book.addSheet('Sheet1');
        table = sheet.addTable('Table1', 0, 5, 0, 3);
    });

    it('name and setName manage the table name', () => {
        shouldThrow(table.name, {});

        expect(table.name()).toBe('Table1');

        shouldThrow(table.setName, table, 1);
        shouldThrow(table.setName, {}, 'NewName');

        expect(table.setName('NewName')).toBe(table);
        expect(table.name()).toBe('NewName');
    });

    it('ref and setRef manage the table ref', () => {
        shouldThrow(table.ref, {});

        expect(typeof table.ref()).toBe('string');

        shouldThrow(table.setRef, table, 1);
        shouldThrow(table.setRef, {}, 'A1:D6');

        expect(table.setRef('A1:D10')).toBe(table);
    });

    it('autoFilter returns an AutoFilter instance', () => {
        shouldThrow(table.autoFilter, {});

        expect(table.autoFilter()).toBeInstanceOf(xl.AutoFilter);
    });

    it('style and setStyle manage the table style', () => {
        shouldThrow(table.style, {});

        expect(typeof table.style()).toBe('number');

        shouldThrow(table.setStyle, table, 'a');
        shouldThrow(table.setStyle, {}, xl.TABLESTYLE_LIGHT1);

        expect(table.setStyle(xl.TABLESTYLE_LIGHT1)).toBe(table);
        expect(table.style()).toBe(xl.TABLESTYLE_LIGHT1);
    });

    it('showRowStripes and setShowRowStripes manage row stripes', () => {
        shouldThrow(table.showRowStripes, {});

        expect(typeof table.showRowStripes()).toBe('boolean');

        shouldThrow(table.setShowRowStripes, table, 'a');
        shouldThrow(table.setShowRowStripes, {}, true);

        expect(table.setShowRowStripes(false)).toBe(table);
        expect(table.showRowStripes()).toBe(false);

        table.setShowRowStripes(true);
        expect(table.showRowStripes()).toBe(true);
    });

    it('showColumnStripes and setShowColumnStripes manage column stripes', () => {
        shouldThrow(table.showColumnStripes, {});

        expect(typeof table.showColumnStripes()).toBe('boolean');

        shouldThrow(table.setShowColumnStripes, table, 'a');
        shouldThrow(table.setShowColumnStripes, {}, true);

        expect(table.setShowColumnStripes(true)).toBe(table);
        expect(table.showColumnStripes()).toBe(true);

        table.setShowColumnStripes(false);
        expect(table.showColumnStripes()).toBe(false);
    });

    it('showFirstColumn and setShowFirstColumn manage first column highlighting', () => {
        shouldThrow(table.showFirstColumn, {});

        expect(typeof table.showFirstColumn()).toBe('boolean');

        shouldThrow(table.setShowFirstColumn, table, 'a');
        shouldThrow(table.setShowFirstColumn, {}, true);

        expect(table.setShowFirstColumn(true)).toBe(table);
        expect(table.showFirstColumn()).toBe(true);

        table.setShowFirstColumn(false);
        expect(table.showFirstColumn()).toBe(false);
    });

    it('showLastColumn and setShowLastColumn manage last column highlighting', () => {
        shouldThrow(table.showLastColumn, {});

        expect(typeof table.showLastColumn()).toBe('boolean');

        shouldThrow(table.setShowLastColumn, table, 'a');
        shouldThrow(table.setShowLastColumn, {}, true);

        expect(table.setShowLastColumn(true)).toBe(table);
        expect(table.showLastColumn()).toBe(true);

        table.setShowLastColumn(false);
        expect(table.showLastColumn()).toBe(false);
    });

    it('columnSize returns the number of columns', () => {
        shouldThrow(table.columnSize, {});

        expect(table.columnSize()).toBe(4);
    });

    it('columnName and setColumnName manage column names', () => {
        shouldThrow(table.columnName, table, 'a');
        shouldThrow(table.columnName, {}, 0);

        expect(typeof table.columnName(0)).toBe('string');

        shouldThrow(table.setColumnName, table, 'a', 'Name');
        shouldThrow(table.setColumnName, table, 0, 1);
        shouldThrow(table.setColumnName, {}, 0, 'Name');

        expect(table.setColumnName(0, 'MyColumn')).toBe(table);
        expect(table.columnName(0)).toBe('MyColumn');
    });
});
